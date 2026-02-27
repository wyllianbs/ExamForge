"""
ExamForge.py — Sorteio de questões a partir de banco .tex configurado via .ods

Leitura de arquivo .ods com estrutura de grupos de questões.
Cada planilha contém: Type, Group, File, n, g, k, Sum

Processamento de arquivos .tex para contagem e extração de questões.
Sorteio com controle de repetição via logs de etapas anteriores.
"""

from __future__ import annotations

import os
import random
import re
import sys
from collections import defaultdict
from dataclasses import dataclass, field
from datetime import date
from typing import Any

from odf.opendocument import load
from odf.table import Table, TableRow, TableCell
from odf.text import P



@dataclass
class Question:
    """Representa uma questão extraída de um arquivo .tex."""
    number: int
    id: str
    line_start: int
    line_end: int
    content: str


@dataclass
class GroupMeta:
    """Metadados de posição de um arquivo dentro de seu grupo."""
    id: int
    sequence: int
    end: int


@dataclass
class FileEntry:
    """Questões extraídas de um único arquivo .tex."""
    type: str
    file: str
    group: GroupMeta
    n: int
    k: int
    questions: list[Question] = field(default_factory=list)

    @property
    def is_single_file_group(self) -> bool:
        return self.group.sequence == 0 and self.group.end == 0


@dataclass
class GroupConfig:
    """Configuração de um grupo lida da planilha ODS."""
    type: str
    group: int | None
    files: list[str]
    n: list[int]
    g: int
    k: int


@dataclass
class DrawRecord:
    """Registro de um sorteio anterior (lido de arquivo .log)."""
    prova: str
    ano: int
    semestre: int
    ids: list[str]

    @property
    def filename(self) -> str:
        return f"{self.prova}-{self.ano}-{self.semestre}.log"


@dataclass
class DrawnQuestion:
    """Uma questão efetivamente sorteada."""
    type: str
    group: int
    file: str
    id: str
    content: str


class Printer:
    """
    Utilitário de formatação padronizada no terminal.

    Centraliza larguras, ícones e estilos de bloco para garantir
    saída visualmente uniforme em todo o programa.
    """

    W      = 75             # largura de todas as réguas
    _THIN  = "─" * W        # régua leve  (seções)
    _THICK = "═" * W        # régua forte (resumo final)

    OK   = "✔"
    ERR  = "✘"
    WARN = "⚠"
    INFO = "ℹ"
    ITEM = "•"
    ARR  = "→"

    # ── primitivos ────────────────────────────────────────────────

    @classmethod
    def blank(cls) -> None:
        print()

    @classmethod
    def rule(cls) -> None:
        print(cls._THIN)

    # ── cabeçalho de seção ────────────────────────────────────────

    @classmethod
    def section(cls, title: str, detail: str = "") -> None:
        """
        Título embutido na régua:
            \n── Título da Seção ──────────────────────────────────────
              detalhe opcional (sem régua inferior)
        """
        label = f"── {title} "
        bar   = cls._THIN[len(label):]
        print()
        print(f"{label}{bar}")
        if detail:
            print(f"  {detail}")

    # ── mensagens de status ───────────────────────────────────────

    @classmethod
    def ok(cls, msg: str) -> None:
        print(f"  {cls.OK}  {msg}")

    @classmethod
    def err(cls, msg: str) -> None:
        print(f"  {cls.ERR}  {msg}")

    @classmethod
    def warn(cls, msg: str) -> None:
        print(f"  {cls.WARN}  {msg}")

    @classmethod
    def info(cls, msg: str) -> None:
        print(f"  {cls.INFO}  {msg}")

    @classmethod
    def item(cls, msg: str, indent: int = 4) -> None:
        print(f"{' ' * indent}{cls.ITEM}  {msg}")

    @classmethod
    def plain(cls, msg: str, indent: int = 2) -> None:
        print(f"{' ' * indent}{msg}")

    # ── bloco de resumo final ─────────────────────────────────────

    @classmethod
    def summary(cls, title: str, rows: list[tuple[str, str]]) -> None:
        """
        Bloco com réguas duplas ═:
            \n════════════════════════════════════════════════════════════
              Título
            ────────────────────────────────────────────────────────────
              Chave  →  Valor
            ════════════════════════════════════════════════════════════\n
        """
        max_k = max(len(k) for k, _ in rows)
        print()
        print(cls._THICK)
        print(f"  {title}")
        print(cls._THIN)
        for k, v in rows:
            print(f"  {k.ljust(max_k)}  {cls.ARR}  {v}")
        print(cls._THICK)
        print()


class OdsReader:
    """Lê e interpreta arquivos .ods."""

    def __init__(self, filepath: str) -> None:
        self._filepath = filepath

    def sheet_names(self) -> list[str]:
        """Retorna os nomes de todas as planilhas do arquivo."""
        doc = load(self._filepath)
        return [s.getAttribute("name") for s in doc.getElementsByType(Table)]

    def read_sheet(self, sheet_name: str) -> list[list[str | None]]:
        """
        Lê uma planilha e retorna lista de linhas (cada linha é uma lista de
        valores). Linhas completamente vazias são descartadas.
        """
        doc = load(self._filepath)
        sheets = {
            s.getAttribute("name"): s
            for s in doc.getElementsByType(Table)
        }

        if sheet_name not in sheets:
            available = list(sheets.keys())
            raise ValueError(
                f"Planilha '{sheet_name}' não encontrada. "
                f"Disponíveis: {available}"
            )

        result: list[list[str | None]] = []

        for row in sheets[sheet_name].getElementsByType(TableRow):
            values = self._parse_row(row)
            if any(v is not None for v in values):
                result.append(values)

        return result

    @staticmethod
    def _parse_row(row: Any) -> list[str | None]:
        values: list[str | None] = []

        for cell in row.getElementsByType(TableCell):
            repeated = cell.getAttribute("numbercolumnsrepeated")
            paragraphs = cell.getElementsByType(P)
            value: str | None = (
                str(paragraphs[0]).strip() if paragraphs else None
            )
            count = int(repeated) if repeated else 1
            values.extend([value] * count)

        # Remove colunas extras vazias no final
        while values and values[-1] is None:
            values.pop()

        return values



class GroupParser:
    """Converte linhas brutas do ODS em GroupConfig."""

    _HEADER_IGNORE = {"type"}

    def parse(self, raw_rows: list[list[str | None]]) -> list[GroupConfig]:
        """
        Converte as linhas brutas em uma lista de GroupConfig.

        Detecção da primeira linha de um novo grupo:
        - Formato completo: g e k ambos presentes e g > 0.
        - Formato parcial: g > 0 e k ausente → assume k = 1.
        - Linha de continuação: g == 0 ou g ausente (sem k).
        """
        groups: list[GroupConfig] = []
        current: GroupConfig | None = None

        for raw in raw_rows:
            row = list(raw) + [None] * (7 - len(raw))

            type_val  = row[0]
            group_val = self._to_int(row[1])
            file_val  = row[2]
            n_val     = self._to_int(row[3])
            g_val     = self._to_int(row[4])
            k_val     = self._to_int(row[5])

            if not type_val or str(type_val).lower() in self._HEADER_IGNORE:
                continue

            if g_val is not None and g_val > 0:
                if current is not None:
                    groups.append(current)

                current = GroupConfig(
                    type=str(type_val),
                    group=group_val,
                    files=[str(file_val)] if file_val else [],
                    n=[n_val] if n_val is not None else [],
                    g=g_val,
                    k=k_val if k_val is not None else 0,  # 0 = não definido
                )
            elif current is not None and file_val:
                current.files.append(str(file_val))
                if n_val is not None:
                    current.n.append(n_val)

        if current is not None:
            groups.append(current)

        return groups

    @staticmethod
    def _to_int(value: Any) -> int | None:
        try:
            return int(float(value)) if value is not None else None
        except (ValueError, TypeError):
            return None

    @staticmethod
    def unique_types(groups: list[GroupConfig]) -> list[str]:
        """Retorna os valores únicos de 'type', preservando a ordem."""
        seen: set[str] = set()
        result: list[str] = []
        for grp in groups:
            if grp.type not in seen:
                seen.add(grp.type)
                result.append(grp.type)
        return result



class FileManager:
    """Gerencia caminhos e verificação de arquivos .tex."""

    def build_type_dir_map(
        self,
        dirs: list[str],
        types: list[str],
    ) -> dict[str, str]:
        """
        Constrói mapeamento type → diretório.

        Se houver apenas 1 diretório, todos os types o usam.
        Com N diretórios, faz match case-insensitive pelo nome da pasta;
        types sem match recebem './<type>' como fallback.
        """
        norm_dirs = [os.path.normpath(d) for d in dirs]

        if len(norm_dirs) == 1:
            return {t: norm_dirs[0] for t in types}

        dir_by_name = {os.path.basename(d).upper(): d for d in norm_dirs}
        mapping: dict[str, str] = {}

        for t in types:
            if t.upper() in dir_by_name:
                mapping[t] = dir_by_name[t.upper()]
            else:
                fallback = os.path.normpath(f"./{t}")
                mapping[t] = fallback
                Printer.warn(
                    f"Type '{t}' sem diretório correspondente."
                    f" Usando fallback: '{fallback}'"
                )

        return mapping

    def apply_paths(
        self,
        groups: list[GroupConfig],
        type_dir_map: dict[str, str],
    ) -> None:
        """Prefixa cada nome de arquivo com o diretório do seu type."""
        for grp in groups:
            base = type_dir_map.get(grp.type, f"./{grp.type}")
            grp.files = [os.path.join(base, f) for f in grp.files]

    def find_missing(self, groups: list[GroupConfig]) -> list[str]:
        """Retorna caminhos de arquivos não encontrados no sistema."""
        missing: list[str] = []
        for grp in groups:
            for filepath in grp.files:
                if not os.path.isfile(filepath):
                    missing.append(filepath)
        return missing

    def report_missing(
        self,
        missing: list[str],
        type_dir_map: dict[str, str],
    ) -> None:
        """Exibe o resultado da verificação de arquivos."""
        dirs_str = ", ".join(type_dir_map.values())
        Printer.section("Verificação de Arquivos", f"Diretórios: {dirs_str}")
        if not missing:
            Printer.ok("Todos os arquivos foram encontrados.")
        else:
            Printer.err(f"{len(missing)} arquivo(s) não encontrado(s):")
            Printer.blank()
            for f in missing:
                Printer.item(f)
        Printer.blank()



class TexParser:
    """Extrai questões de arquivos .tex usando rastreamento de chaves."""

    # { seguido de % e um ID (com possíveis espaços)
    _RE_ID = re.compile(r'\{\s*%\s*(\S+)')

    def extract(self, filepath: str) -> list[Question]:
        """
        Lê um arquivo .tex e extrai todas as questões.

        Usa pilha de blocos para rastrear aninhamento de chaves.
        \\rtask marca o bloco mais externo como questão.
        """
        try:
            with open(filepath, encoding="utf-8", errors="ignore") as fh:
                lines = fh.readlines()
        except OSError as exc:
            print(f"  ✘ Erro ao ler arquivo {filepath}: {exc}")
            return []

        questions: list[Question] = []
        stack: list[dict[str, Any]] = []

        acc_lines: list[str] = []
        start_line = 0
        current_id = ""
        inside = False

        for i, line in enumerate(lines):
            num = i + 1
            has_rtask = "\\rtask" in line
            id_match = self._RE_ID.search(line)
            opens = line.count("{")
            closes = line.count("}")

            # Bloco malformado: novo ID com pilha não vazia
            if opens > 0 and id_match and stack:
                stack.clear()
                inside = False
                acc_lines = []

            for _ in range(opens):
                block: dict[str, Any] = {
                    "line_start": num,
                    "has_rtask": False,
                    "id": None,
                }
                if id_match:
                    block["id"] = id_match.group(1)
                    id_match = None
                stack.append(block)

            if stack and has_rtask and not stack[0]["has_rtask"]:
                stack[0]["has_rtask"] = True

            if stack and stack[0]["has_rtask"]:
                if not inside:
                    inside = True
                    start_line = int(stack[0]["line_start"])
                    current_id = str(stack[0].get("id") or "")
                    acc_lines = [
                        lines[j] for j in range(start_line - 1, num)
                    ]
                else:
                    acc_lines.append(line)
            elif inside and not stack:
                pass
            elif not inside:
                pass
            else:
                acc_lines.append(line)

            for _ in range(closes):
                if stack:
                    closed = stack.pop()
                    if closed["has_rtask"] and not stack:
                        questions.append(Question(
                            number=len(questions) + 1,
                            id=current_id,
                            line_start=start_line,
                            line_end=num,
                            content="".join(acc_lines),
                        ))
                        inside = False
                        acc_lines = []

        if inside and stack:
            print(
                f"  ⚠  Questão não fechada em {filepath}, "
                f"linha {start_line}, ID: {current_id}"
            )

        return questions

    def verify_counts(
        self,
        groups: list[GroupConfig],
    ) -> tuple[bool, list[dict[str, Any]]]:
        """
        Compara a contagem de questões do ODS com o que há nos .tex.

        Retorna (tudo_ok, lista_de_entradas_brutas).
        """
        all_entries: list[dict[str, Any]] = []
        divergences: list[dict[str, Any]] = []
        cache: dict[str, list[Question]] = {}

        for grp in groups:
            for filepath, n_ods in zip(grp.files, grp.n):
                if filepath not in cache:
                    cache[filepath] = self.extract(filepath)

                questions = cache[filepath]
                n_tex = len(questions)

                entry: dict[str, Any] = {
                    "type":      grp.type,
                    "group":     grp.group,
                    "file":      filepath,
                    "n_ods":     n_ods,
                    "n_tex":     n_tex,
                    "questions": questions,
                }
                all_entries.append(entry)

                if n_tex != n_ods:
                    divergences.append(entry)

        if divergences:
            Printer.section("Divergência na Contagem de Questões")
            Printer.blank()
            for div in divergences:
                kv_rows = [
                    ("Arquivo", div['file']),
                    ("Type",    div['type']),
                    ("Group",   str(div['group'])),
                    ("n (ODS)", str(div['n_ods'])),
                    ("n (.tex)", str(div['n_tex'])),
                ]
                max_k = max(len(k) for k, _ in kv_rows)
                for k, v in kv_rows:
                    Printer.plain(f"{k.ljust(max_k)}  {Printer.ARR}  {v}")
                Printer.blank()
            Printer.err(f"Total de divergências: {len(divergences)}")
            Printer.blank()
            return False, all_entries

        return True, all_entries



class EntryBuilder:
    """Constrói FileEntry a partir dos dados brutos de extração."""

    def build(
        self,
        raw_entries: list[dict[str, Any]],
        groups: list[GroupConfig],
    ) -> list[FileEntry]:
        """Constrói a lista de FileEntry, numerando as questões."""
        file_meta = self._build_file_meta(groups)
        result: list[FileEntry] = []
        seen: set[str] = set()

        for raw in raw_entries:
            filepath: str = raw["file"]
            if filepath in seen:
                continue
            seen.add(filepath)

            grp_id, k, seq, end = file_meta.get(filepath, (0, 1, 0, 0))

            numbered = [
                Question(
                    number=idx,
                    id=q.id,
                    line_start=q.line_start,
                    line_end=q.line_end,
                    content=q.content,
                )
                for idx, q in enumerate(raw["questions"], start=1)
            ]

            result.append(FileEntry(
                type=raw["type"],
                file=filepath,
                group=GroupMeta(id=grp_id, sequence=seq, end=end),
                n=raw["n_tex"],
                k=k,
                questions=numbered,
            ))

        return result

    @staticmethod
    def _build_file_meta(
        groups: list[GroupConfig],
    ) -> dict[str, tuple[int, int, int, int]]:
        """Mapeia caminho de arquivo → (group_id, k, sequence, end)."""
        meta: dict[str, tuple[int, int, int, int]] = {}
        for grp in groups:
            end = len(grp.files) - 1
            for seq, fname in enumerate(grp.files):
                meta[fname] = (grp.group or 0, grp.k, seq, end)
        return meta

    @staticmethod
    def print_summary(entries: list[FileEntry]) -> None:
        """Imprime resumo tabular das questões extraídas."""
        total_q = sum(e.n for e in entries)
        sep = "═" * 60
        print(f"\n{sep}")
        print("  Estrutura de questões extraídas")
        print(sep)
        print(f"  Total de arquivos : {len(entries)}")
        print(f"  Total de questões : {total_q}")
        print(f"{sep}\n")

        if not entries:
            return

        prefixes = [f"{e.type} │ {e.file}" for e in entries]
        max_pre = max(len(p) for p in prefixes)
        max_n   = max(len(str(e.n)) for e in entries)
        soma    = 0

        for prefix, entry in zip(prefixes, entries):
            soma += entry.n
            dots = "." * max(0, max_pre - len(prefix) - 1)
            sep_char = f" {dots} " if dots else " "
            print(f"  {prefix}{sep_char}: {entry.n:>{max_n}}")

        total_width = max_pre + 3 + max_n + 2
        print(f"  {'─' * total_width}")
        label = "Total"
        dots = "." * max(0, max_pre - len(label) - 1)
        sep_char = f" {dots} " if dots else " "
        print(f"  {label}{sep_char}: {soma:>{max_n}}")
        print()



class LogManager:
    """Lê e grava logs de sorteios (.log)."""

    _RE_LOG_NAME = re.compile(r'^(.+)-(\d{4})-(\d+)\.log$')

    def read_directory(self, log_dir: str) -> list[DrawRecord]:
        """
        Lê todos os .log de um diretório e retorna DrawRecords
        ordenados por (ano, semestre) crescente.
        """
        records: list[DrawRecord] = []

        try:
            entries = os.listdir(log_dir)
        except OSError as exc:
            print(f"  ⚠  Erro ao listar '{log_dir}': {exc}")
            return records

        for filename in sorted(entries):
            match = self._RE_LOG_NAME.match(filename)
            if not match:
                continue

            filepath = os.path.join(log_dir, filename)
            try:
                with open(filepath, encoding="utf-8") as fh:
                    ids = [ln.strip() for ln in fh if ln.strip()]
            except OSError as exc:
                print(f"  ⚠  Erro ao ler '{filepath}': {exc}")
                continue

            records.append(DrawRecord(
                prova=match.group(1),
                ano=int(match.group(2)),
                semestre=int(match.group(3)),
                ids=ids,
            ))

        records.sort(key=lambda r: (r.ano, r.semestre))
        return records

    def load_previous(self, sheet_name: str) -> list[DrawRecord]:
        """
        Carrega os logs de sorteios anteriores para a planilha informada.
        Oferece criar o diretório se não existir.
        """
        log_dir = os.path.join("logs", sheet_name)
        Printer.section("Histórico de Sorteios", f"Diretório: {log_dir}")

        if not os.path.isdir(log_dir):
            Printer.info(f"Diretório '{log_dir}' não encontrado.")
            answer = input(
                f"     Deseja criá-lo recursivamente? [S/n]: "
            ).strip().lower()

            if answer != "n":
                try:
                    os.makedirs(log_dir, exist_ok=True)
                    Printer.ok(f"Diretório '{log_dir}' criado.")
                except OSError as exc:
                    Printer.err(f"Erro ao criar '{log_dir}': {exc}")
            else:
                Printer.plain("Diretório não criado.")

            Printer.plain("Nenhum histórico de sorteios anteriores será usado.")
            Printer.blank()
            return []

        records = self.read_directory(log_dir)

        if not records:
            Printer.info(f"Nenhum log encontrado em '{log_dir}'.")
            Printer.blank()
            return []

        total_ids = sum(len(r.ids) for r in records)
        Printer.ok(f"{len(records)} log(s) encontrado(s)  ·  {total_ids} IDs no total")
        Printer.blank()
        for rec in records:
            Printer.item(f"{rec.filename:35s}  ({len(rec.ids):>4d} IDs)")

        Printer.blank()
        return records

    def save(self, drawn_ids: list[str], filepath: str) -> None:
        """Salva IDs sorteados em arquivo .log (um por linha)."""
        try:
            log_dir = os.path.dirname(filepath)
            if log_dir:
                os.makedirs(log_dir, exist_ok=True)
            with open(filepath, "w", encoding="utf-8") as fh:
                fh.writelines(f"{qid}\n" for qid in drawn_ids)
            Printer.ok(f"Log salvo: {filepath}  ({len(drawn_ids)} IDs)")
        except OSError as exc:
            Printer.err(f"Erro ao salvar log '{filepath}': {exc}")




class QuestionDrawer:
    """
    Sorteia questões respeitando:
    - Prioridade para questões inéditas.
    - Round-robin entre arquivos de um mesmo grupo.
    - Reutilização das mais antigas quando necessário.
    """

    def draw(
        self,
        entries: list[FileEntry],
        previous: list[DrawRecord],
    ) -> tuple[list[DrawnQuestion], list[str]]:
        """
        Retorna (questões_sorteadas, ids_sorteados).
        """
        drawn: list[DrawnQuestion] = []
        drawn_ids: list[str] = []

        all_prev_ids = self._all_previous_ids(previous)
        group_map = self._build_group_map(entries)

        processed: set[tuple[str, int]] = set()

        for entry in entries:
            key = (entry.type, entry.group.id)
            if key in processed:
                continue
            processed.add(key)

            group_entries = group_map[key]
            k = entry.k

            group_ids = self._collect_group_ids(group_entries)
            inedit_ids = group_ids - all_prev_ids
            used_ids   = group_ids & all_prev_ids

            reuse_ordered: list[str] = []

            if len(inedit_ids) < k:
                reuse_ordered = self._handle_shortage(
                    entry, group_entries, group_ids, inedit_ids,
                    used_ids, previous, k,
                )
                if reuse_ordered is None:          # usuário recusou
                    continue

            if entry.is_single_file_group:
                selected = self._select_single(
                    entry, k, inedit_ids, reuse_ordered
                )
            else:
                selected = self._select_round_robin(
                    group_entries, k, inedit_ids, reuse_ordered
                )

            for q in selected:
                source = self._find_source(q.id, group_entries)
                drawn.append(DrawnQuestion(
                    type=entry.type,
                    group=entry.group.id,
                    file=source,
                    id=q.id,
                    content=q.content,
                ))
                drawn_ids.append(q.id)

        return drawn, drawn_ids

    # ── helpers privados ──────────────────────────────────────────────

    @staticmethod
    def _all_previous_ids(records: list[DrawRecord]) -> set[str]:
        ids: set[str] = set()
        for rec in records:
            ids.update(rec.ids)
        return ids

    @staticmethod
    def _build_group_map(
        entries: list[FileEntry],
    ) -> dict[tuple[str, int], list[FileEntry]]:
        mapping: dict[tuple[str, int], list[FileEntry]] = defaultdict(list)
        for e in entries:
            mapping[(e.type, e.group.id)].append(e)
        for key in mapping:
            mapping[key].sort(key=lambda e: e.group.sequence)
        return mapping

    @staticmethod
    def _collect_group_ids(entries: list[FileEntry]) -> set[str]:
        return {q.id for e in entries for q in e.questions}

    @staticmethod
    def _find_source(qid: str, entries: list[FileEntry]) -> str:
        for e in entries:
            if any(q.id == qid for q in e.questions):
                return e.file
        return ""

    def _handle_shortage(
        self,
        entry: FileEntry,
        group_entries: list[FileEntry],
        group_ids: set[str],
        inedit_ids: set[str],
        used_ids: set[str],
        previous: list[DrawRecord],
        k: int,
    ) -> list[str] | None:
        """
        Lida com a falta de questões inéditas.
        Retorna a lista ordenada de IDs reutilizáveis,
        ou None se o usuário recusar a reutilização.
        """
        needed = k - len(inedit_ids)
        group_files = [e.file for e in group_entries]

        if not previous or not used_ids:
            self._warn_shortage(entry, inedit_ids, k, group_files, used_ids)
            return []

        etapas, reuse_ids = self._find_reuse(previous, group_ids, needed)

        self._print_reuse_prompt(entry, inedit_ids, group_ids,
                                  group_files, needed, etapas)

        answer = input(
            "\n     Deseja prosseguir com a reutilização? [S/n]: "
        ).strip().lower()

        if answer == "n":
            Printer.err(
                f"Grupo ({entry.type}, id={entry.group.id}) ignorado pelo usuário."
            )
            Printer.blank()
            return None

        Printer.blank()
        return reuse_ids

    @staticmethod
    def _warn_shortage(
        entry: FileEntry,
        inedit_ids: set[str],
        k: int,
        group_files: list[str],
        used_ids: set[str],
    ) -> None:
        Printer.warn(
            f"Grupo ({entry.type}, id={entry.group.id}): "
            f"apenas {len(inedit_ids)} questão(ões) inédita(s), k={k}."
        )
        Printer.plain("Arquivos do grupo:", indent=5)
        for gf in group_files:
            Printer.item(gf, indent=6)
        if len(inedit_ids) == 0 and not used_ids:
            Printer.err("Nenhuma questão disponível. Grupo ignorado.")
            Printer.blank()

    @staticmethod
    def _print_reuse_prompt(
        entry: FileEntry,
        inedit_ids: set[str],
        group_ids: set[str],
        group_files: list[str],
        needed: int,
        etapas: list[DrawRecord],
    ) -> None:
        if len(inedit_ids) == 0:
            Printer.warn(
                f"Grupo ({entry.type}, id={entry.group.id}): "
                f"todas as {len(group_ids)} questões já foram sorteadas."
            )
        else:
            Printer.warn(
                f"Grupo ({entry.type}, id={entry.group.id}): "
                f"apenas {len(inedit_ids)} de {len(group_ids)} questões são inéditas."
            )
        Printer.plain("Arquivos do grupo:", indent=5)
        for gf in group_files:
            Printer.item(gf, indent=6)
        Printer.blank()
        Printer.plain(
            f"Para atingir k={entry.k}, serão reutilizadas "
            f"{needed} questão(ões) das etapas mais antigas:",
            indent=5,
        )
        for etapa in etapas:
            Printer.item(
                f"{etapa.filename}  (Ano: {etapa.ano}, Sem.: {etapa.semestre})",
                indent=6,
            )

    @staticmethod
    def _find_reuse(
        previous: list[DrawRecord],
        group_ids: set[str],
        needed: int,
    ) -> tuple[list[DrawRecord], list[str]]:
        """Busca IDs reutilizáveis nas etapas mais antigas."""
        etapas: list[DrawRecord] = []
        reuse_ids: list[str] = []
        collected = 0

        for rec in previous:
            ids_in_group = [qid for qid in rec.ids if qid in group_ids]
            if ids_in_group:
                etapas.append(rec)
                reuse_ids.extend(ids_in_group)
                collected += len(ids_in_group)
                if collected >= needed:
                    break

        return etapas, reuse_ids

    @staticmethod
    def _select_single(
        entry: FileEntry,
        k: int,
        inedit_ids: set[str],
        reuse_ordered: list[str],
    ) -> list[Question]:
        inedit = [q for q in entry.questions if q.id in inedit_ids]
        random.shuffle(inedit)
        selected = inedit[:k]

        if len(selected) < k:
            reuse_map = {
                q.id: q for q in entry.questions if q.id not in inedit_ids
            }
            sel_ids = {q.id for q in selected}
            for qid in reuse_ordered:
                if len(selected) >= k:
                    break
                if qid in reuse_map and qid not in sel_ids:
                    selected.append(reuse_map[qid])
                    sel_ids.add(qid)

        return selected

    @staticmethod
    def _select_round_robin(
        entries: list[FileEntry],
        k: int,
        inedit_ids: set[str],
        reuse_ordered: list[str],
    ) -> list[Question]:
        num = len(entries)
        pools = [
            random.sample(
                [q for q in e.questions if q.id in inedit_ids],
                len([q for q in e.questions if q.id in inedit_ids])
            )
            for e in entries
        ]

        selected: list[Question] = []
        positions = [0] * num
        exhausted = [False] * num
        idx = 0

        while len(selected) < k:
            if all(exhausted):
                break
            attempts = 0
            while exhausted[idx] and attempts < num:
                idx = (idx + 1) % num
                attempts += 1
            if attempts >= num:
                break

            pos = positions[idx]
            pool = pools[idx]
            if pos < len(pool):
                selected.append(pool[pos])
                positions[idx] = pos + 1
            else:
                exhausted[idx] = True

            idx = (idx + 1) % num

        if len(selected) < k:
            sel_ids = {q.id for q in selected}
            reuse_map = {
                q.id: q
                for e in entries
                for q in e.questions
                if q.id not in inedit_ids
            }
            for qid in reuse_ordered:
                if len(selected) >= k:
                    break
                if qid in reuse_map and qid not in sel_ids:
                    selected.append(reuse_map[qid])
                    sel_ids.add(qid)

        return selected



class TexWriter:
    """Grava questões sorteadas em arquivo .tex."""

    def save(self, questions: list[DrawnQuestion], filepath: str) -> None:
        """Salva questões separadas por 3 linhas em branco."""
        try:
            with open(filepath, "w", encoding="utf-8") as fh:
                for idx, q in enumerate(questions):
                    fh.write(q.content)
                    if idx < len(questions) - 1:
                        fh.write("\n\n\n")
            Printer.ok(f"Arquivo salvo: {filepath}  ({len(questions)} questões)")
        except OSError as exc:
            Printer.err(f"Erro ao salvar '{filepath}': {exc}")



class UserInterface:
    """Responsável por toda a interação com o usuário via terminal."""

    VALID_SHEETS  = ["P1", "P1C2", "P2", "P2C2", "Rec"]
    DEFAULT_SHEET = "P1"
    DEFAULT_ODS   = "db.ods"

    def ask_filepath(self) -> str:
        path = input(
            f"\nInforme o caminho do arquivo ODS\n"
            f"[Enter = {self.DEFAULT_ODS}]: "
        ).strip()
        return path or self.DEFAULT_ODS

    def ask_sheet(self) -> str:
        options = ", ".join(self.VALID_SHEETS)
        choice = input(
            f"\nQual configuração deseja ler?\n"
            f"Opções: {options}\n"
            f"[Enter = {self.DEFAULT_SHEET}]: "
        ).strip()

        if not choice:
            return self.DEFAULT_SHEET
        if choice in self.VALID_SHEETS:
            return choice

        print(f"  Opção inválida: '{choice}'. Usando '{self.DEFAULT_SHEET}'.")
        return self.DEFAULT_SHEET

    def ask_dirs(self, sheet_name: str, types: list[str]) -> list[str]:
        default = types[0] if len(types) == 1 else " ".join(types)
        raw = input(
            f"\nTypes encontrados em '{sheet_name}': {', '.join(types)}\n"
            f"Informe o(s) diretório(s) dos arquivos (separados por espaço)\n"
            f"[Enter = {default}]: "
        ).strip()
        return (raw or default).split()

    def ask_output_filepath(self, sheet_name: str) -> str:
        default = f"{sheet_name}.tex"
        path = input(
            f"\nInforme o caminho do arquivo de saída (.tex)\n"
            f"[Enter = {default}]: "
        ).strip()
        return path or default

    DEFAULT_RANDOM_COUNT = 50

    def ask_random_count(self, n_inedit: int, n_total: int) -> int:
        """
        Pergunta ao usuário quantas questões sortear quando 'k' não
        está preenchido na planilha. Exibe o contexto de questões
        inéditas disponíveis. Default: 50.
        """
        default = self.DEFAULT_RANDOM_COUNT
        Printer.section("Modo Aleatório — Coluna k Não Preenchida")
        Printer.blank()
        rows = [
            ("Questões inéditas disponíveis",       str(n_inedit)),
            ("Total no banco (incl. já sorteadas)", str(n_total)),
        ]
        max_k = max(len(k) for k, _ in rows)
        for k, v in rows:
            Printer.plain(f"{k.ljust(max_k)}  {Printer.ARR}  {v}")
        Printer.blank()

        raw = input(
            f"  Quantas questões deseja sortear ao todo?\n"
            f"  [Enter = {default}]: "
        ).strip()

        if not raw:
            return default
        try:
            val = int(raw)
            if val < 1:
                Printer.warn(f"Valor deve ser ≥ 1. Usando {default}.")
                return default
            return val
        except ValueError:
            Printer.warn(f"Valor inválido. Usando {default}.")
            return default

    def ask_log_filepath(self, output_filepath: str) -> str:
        today = date.today()
        semester = 1 if today.month <= 7 else 2
        base = os.path.splitext(os.path.basename(output_filepath))[0]
        default = os.path.join("logs", base, f"{base}-{today.year}-{semester}.log")

        path = input(
            f"\nInforme o caminho do arquivo de log\n"
            f"[Enter = {default}]: "
        ).strip()
        return path or default


class ExamForge:
    """Orquestra todo o fluxo do ExamForge."""

    def __init__(self) -> None:
        self._ui          = UserInterface()
        self._file_mgr    = FileManager()
        self._parser      = GroupParser()
        self._tex_parser  = TexParser()
        self._builder     = EntryBuilder()
        self._log_mgr     = LogManager()
        self._drawer      = QuestionDrawer()
        self._tex_writer  = TexWriter()

    @staticmethod
    def _distribute_k(
        group_map: dict[tuple, list[FileEntry]],
        group_inedit: dict[tuple, int],
        group_total: dict[tuple, int],
        target: int,
    ) -> None:
        """
        Distribui 'target' questões proporcionalmente entre os grupos
        com k não definido (sentinela 0).

        Grupos com peso muito baixo podem receber k=0 e serão ignorados
        no sorteio — isso garante que o total sorteado seja exatamente
        'target' e não seja inflacionado pelo número de grupos.

        Fluxo:
        1. Peso = questões inéditas do grupo (fallback: total).
        2. Alocação proporcional sem mínimo forçado (k pode ser 0).
        3. Ajuste fino para a soma bater exatamente 'target'.
        4. Define entry.k em todos os FileEntry de cada grupo.
        """
        keys = list(group_map.keys())
        n_groups = len(keys)

        if n_groups == 0:
            return

        # Pesos: prioriza inéditas; fallback para total
        weights = [group_inedit[k] for k in keys]
        total_w = sum(weights)
        if total_w == 0:
            weights = [group_total[k] for k in keys]
            total_w = sum(weights)
        if total_w == 0:
            weights = [1] * n_groups
            total_w = n_groups

        # Distribuição proporcional — sem mínimo; grupos pequenos ficam em 0
        raw = [target * w / total_w for w in weights]
        k_vals = [round(r) for r in raw]   # pode resultar em 0

        # Ajuste fino para bater o total exato
        diff = target - sum(k_vals)
        if diff != 0:
            # Ordena pelos resíduos fracionários maiores (para aumentar)
            # ou menores (para diminuir), mas nunca abaixo de 0
            fracs = sorted(
                range(n_groups),
                key=lambda i: raw[i] - int(raw[i]),
                reverse=(diff > 0),
            )
            for i in fracs:
                if diff == 0:
                    break
                if diff > 0:
                    k_vals[i] += 1
                    diff -= 1
                else:
                    if k_vals[i] > 0:   # não deixa negativo
                        k_vals[i] -= 1
                        diff += 1

        Printer.section("Distribuição Automática de k por Grupo")
        Printer.blank()

        skipped = 0
        for key, k in zip(keys, k_vals):
            for entry in group_map[key]:
                entry.k = k
            type_s, grp_id = key
            if k == 0:
                skipped += 1
                Printer.plain(
                    f"Grupo ({type_s}, id={grp_id:>3})  "
                    f"{Printer.ARR}  k = 0  (ignorado)",
                    indent=2,
                )
            else:
                Printer.plain(
                    f"Grupo ({type_s}, id={grp_id:>3})  "
                    f"{Printer.ARR}  k = {k}"
                    f"   inéditas: {group_inedit[key]}   "
                    f"total: {group_total[key]}",
                    indent=2,
                )

        effective = sum(k_vals)
        Printer.blank()
        rows: list[tuple[str, str]] = [("Total a sortear", str(effective))]
        if skipped:
            rows.append(("Grupos ignorados (k=0)", str(skipped)))
        max_k = max(len(r) for r, _ in rows)
        for r, v in rows:
            Printer.plain(f"{r.ljust(max_k)}  {Printer.ARR}  {v}")
        Printer.blank()

    def run(self) -> None:
        # 1. Configuração inicial
        ods_path   = self._ui.ask_filepath()
        sheet_name = self._ui.ask_sheet()

        Printer.section("Lendo Configuração", f"Planilha: {sheet_name}   Arquivo: {ods_path}")

        ods    = OdsReader(ods_path)
        raw    = ods.read_sheet(sheet_name)
        groups = self._parser.parse(raw)
        types  = self._parser.unique_types(groups)

        # 2. Diretórios e caminhos
        dirs         = self._ui.ask_dirs(sheet_name, types)
        type_dir_map = self._file_mgr.build_type_dir_map(dirs, types)

        Printer.section("Mapeamento Type → Diretório")
        for t, d in type_dir_map.items():
            Printer.item(f"{t}  {Printer.ARR}  {d}")

        self._file_mgr.apply_paths(groups, type_dir_map)

        missing = self._file_mgr.find_missing(groups)
        self._file_mgr.report_missing(missing, type_dir_map)

        if missing:
            Printer.err("Execução interrompida: corrija os arquivos ausentes.")
            Printer.blank()
            sys.exit(1)

        # 3. Verificação de contagens
        Printer.section("Verificação de Contagens (ODS vs .tex)")

        ok, raw_entries = self._tex_parser.verify_counts(groups)
        if not ok:
            Printer.err("Execução interrompida: corrija as divergências.")
            Printer.blank()
            sys.exit(1)

        Printer.ok("Todas as contagens coincidem.")

        # 4. Estrutura de questões
        entries = self._builder.build(raw_entries, groups)
        self._builder.print_summary(entries)

        # 5. Histórico de sorteios
        previous = self._log_mgr.load_previous(sheet_name)

        # 5b. Distribuição automática de k quando não definido na planilha
        unset_entries = [e for e in entries if e.k == 0]
        if unset_entries:
            all_prev_ids: set[str] = {
                qid for rec in previous for qid in rec.ids
            }

            # Agrupa por (type, group_id) para calcular pesos
            group_unset_map: dict[tuple, list[FileEntry]] = defaultdict(list)
            for e in unset_entries:
                group_unset_map[(e.type, e.group.id)].append(e)

            group_inedit: dict[tuple, int] = {}
            group_total:  dict[tuple, int] = {}
            for key, grp_entries in group_unset_map.items():
                all_ids = {q.id for e in grp_entries for q in e.questions}
                group_inedit[key] = len(all_ids - all_prev_ids)
                group_total[key]  = len(all_ids)

            total_inedit = sum(group_inedit.values())
            total_all    = sum(group_total.values())

            target = self._ui.ask_random_count(total_inedit, total_all)
            self._distribute_k(
                group_unset_map, group_inedit, group_total, target
            )

        # 6. Caminhos de saída
        output_path = self._ui.ask_output_filepath(sheet_name)
        log_path    = self._ui.ask_log_filepath(output_path)

        # 7. Sorteio
        Printer.section("Sorteando Questões")
        Printer.blank()

        drawn, drawn_ids = self._drawer.draw(entries, previous)

        # 8. Salvamento
        self._tex_writer.save(drawn, output_path)
        self._log_mgr.save(drawn_ids, log_path)

        # 9. Resumo final
        Printer.summary(
            "Resumo do Sorteio",
            [
                ("Questões sorteadas", str(len(drawn))),
                ("Arquivo .tex",       output_path),
                ("Arquivo .log",       log_path),
            ],
        )



def main() -> None:
    ExamForge().run()


if __name__ == "__main__":
    main()
