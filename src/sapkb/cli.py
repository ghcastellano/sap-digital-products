"""
CLI — Typer-based command line interface for SAP Consulting KB
"""
import sys
from pathlib import Path
from typing import Optional

import typer
from rich.console import Console
from rich.table import Table
from rich.panel import Panel
from rich.text import Text

# Add project root to path
sys.path.insert(0, str(Path(__file__).parent.parent.parent.parent))

from config.settings import get_settings

app = typer.Typer(name="sap-cli", help="SAP Digital Products — Consulting Intelligence System")
console = Console()


def _get_store():
    from .store import KnowledgeStore
    settings = get_settings()
    return KnowledgeStore(
        db_path=str(settings.chroma_db_path),
        collection_name=settings.chroma_collection_name,
        embedding_model=settings.embedding_model,
    )


def _get_searcher():
    from .search import HybridSearcher
    settings = get_settings()
    store = _get_store()
    return HybridSearcher(store, str(settings.bm25_index_path), rrf_k=settings.rrf_k)


# ─── INGEST ───

@app.command()
def ingest(
    path: str = typer.Argument(..., help="File or directory path to ingest"),
    project: str = typer.Option("", "--project", "-p", help="Project name"),
    tags: str = typer.Option("", "--tags", "-t", help="Comma-separated tags"),
    glob: str = typer.Option("*.*", "--glob", "-g", help="Glob pattern for directories"),
    rebuild_index: bool = typer.Option(True, "--rebuild/--no-rebuild", help="Rebuild BM25 index after ingestion"),
    visual: bool = typer.Option(False, "--visual", "-v", help="Use Claude Vision for rich slide descriptions (Haiku for text, Sonnet for complex visuals)"),
):
    """Ingest files into the knowledge base."""
    from .ingestion.orchestrator import IngestionOrchestrator

    store = _get_store()
    orchestrator = IngestionOrchestrator(store)
    tag_list = [t.strip() for t in tags.split(",") if t.strip()] if tags else []
    target = Path(path)

    if visual:
        console.print("[bold yellow]Visual ingest enabled[/bold yellow] — Haiku for text slides, Sonnet for complex visuals")

    if target.is_file():
        with console.status(f"Ingesting {target.name}..."):
            result = orchestrator.ingest_file(target, project=project, tags=tag_list, visual=visual)
        _print_ingest_result(result)
    elif target.is_dir():
        files = list(target.glob(glob))
        files = [f for f in files if f.is_file() and not f.name.startswith(".")]
        console.print(f"Found [bold]{len(files)}[/bold] files matching '{glob}'")
        results = []
        for fp in files:
            with console.status(f"Ingesting {fp.name}..."):
                r = orchestrator.ingest_file(fp, project=project, tags=tag_list, visual=visual)
            _print_ingest_result(r)
            results.append(r)
        total = sum(r.chunks_added for r in results)
        console.print(f"\n[bold green]Total: {total} chunks from {len(results)} files[/bold green]")
    else:
        console.print(f"[red]Path not found: {path}[/red]")
        raise typer.Exit(1)

    if rebuild_index:
        with console.status("Rebuilding BM25 index..."):
            searcher = _get_searcher()
            searcher.build_bm25_index()
        console.print("[green]BM25 index rebuilt.[/green]")


def _print_ingest_result(result):
    status = "[green]OK[/green]" if result.chunks_added > 0 else "[yellow]SKIP[/yellow]"
    proj = f" [{result.project}]" if result.project else ""
    console.print(f"  {status} {result.source_file}{proj} → {result.chunks_added} chunks")
    for w in result.warnings:
        console.print(f"    [yellow]⚠ {w}[/yellow]")


# ─── INGEST TEXT ───

@app.command()
def ingest_text(
    text: str = typer.Option(..., "--text", "-x", help="Raw text to ingest"),
    doc_type: str = typer.Option("notes", "--type", help="Document type: notes or transcript"),
    project: str = typer.Option("", "--project", "-p"),
    speaker: str = typer.Option("", "--speaker", "-s"),
    name: str = typer.Option("raw_input", "--name", "-n", help="Source name identifier"),
    tags: str = typer.Option("", "--tags", "-t"),
):
    """Ingest raw text (transcript or notes) into the knowledge base."""
    from .ingestion.orchestrator import IngestionOrchestrator
    from .models import DocumentType

    store = _get_store()
    orchestrator = IngestionOrchestrator(store)
    dtype = DocumentType.TRANSCRIPT if doc_type == "transcript" else DocumentType.NOTES
    tag_list = [t.strip() for t in tags.split(",") if t.strip()] if tags else []

    result = orchestrator.ingest_raw_text(
        text, doc_type=dtype, source_name=name,
        project=project, speaker=speaker, tags=tag_list,
    )
    _print_ingest_result(result)

    with console.status("Rebuilding BM25 index..."):
        searcher = _get_searcher()
        searcher.build_bm25_index()


# ─── SEARCH ───

@app.command()
def search(
    query: str = typer.Argument(..., help="Search query"),
    top_k: int = typer.Option(10, "--top-k", "-k"),
    project: str = typer.Option("", "--project", "-p"),
    doc_type: str = typer.Option("", "--type", help="Filter by document type"),
    show_text: bool = typer.Option(False, "--text", help="Show full chunk text"),
    check_divergences: bool = typer.Option(False, "--divergences", "-d", help="Check for divergences"),
):
    """Search the knowledge base with hybrid search."""
    filters = {}
    if project:
        filters["project"] = project
    if doc_type:
        filters["document_type"] = doc_type

    searcher = _get_searcher()
    with console.status("Searching..."):
        results = searcher.search(query, top_k=top_k, filters=filters or None)

    if not results:
        console.print("[yellow]No results found.[/yellow]")
        return

    table = Table(title=f"Results for: '{query}'", show_lines=True)
    table.add_column("#", style="dim", width=3)
    table.add_column("Source", style="cyan", max_width=40)
    table.add_column("Score", width=7)
    table.add_column("Project", width=15)
    table.add_column("Preview", max_width=80)

    for i, r in enumerate(results, 1):
        m = r.chunk.metadata
        loc = ""
        if m.slide_number is not None:
            loc = f" [Slide {m.slide_number}]"
        elif m.speaker:
            loc = f" [{m.speaker}]"

        preview = r.chunk.text[:150].replace("\n", " ")
        if len(r.chunk.text) > 150:
            preview += "..."

        table.add_row(
            str(i),
            f"{m.source_file}{loc}",
            f"{r.score:.4f}",
            m.project,
            preview,
        )

    console.print(table)

    if show_text:
        for i, r in enumerate(results, 1):
            console.print(Panel(
                r.chunk.text,
                title=f"[{i}] {r.format_source()}",
                border_style="blue",
            ))

    if check_divergences:
        from .divergence import DivergenceDetector
        detector = DivergenceDetector(searcher.store)
        divs = detector.check_query_results(results)
        if divs:
            console.print(f"\n[bold red]⚠ {len(divs)} potential divergence(s) detected:[/bold red]")
            for d in divs:
                console.print(Panel(
                    f"[bold]Source A:[/bold] {d.source_a}\n{d.chunk_a_text[:200]}...\n\n"
                    f"[bold]Source B:[/bold] {d.source_b}\n{d.chunk_b_text[:200]}...\n\n"
                    f"Similarity: {d.similarity:.3f}",
                    title="[red]DIVERGÊNCIA[/red]",
                    border_style="red",
                ))
        else:
            console.print("\n[green]No divergences detected in results.[/green]")


# ─── STATS ───

@app.command()
def stats():
    """Show knowledge base statistics."""
    store = _get_store()
    s = store.get_stats()

    console.print(Panel(
        f"[bold]Total chunks:[/bold] {s['total_chunks']}\n"
        f"[bold]Total documents:[/bold] {s.get('total_documents', 0)}",
        title="Knowledge Base Stats",
        border_style="blue",
    ))

    if s["by_project"]:
        table = Table(title="By Project")
        table.add_column("Project"); table.add_column("Chunks", justify="right")
        for proj, count in sorted(s["by_project"].items()):
            table.add_row(proj, str(count))
        console.print(table)

    if s["by_type"]:
        table = Table(title="By Document Type")
        table.add_column("Type"); table.add_column("Chunks", justify="right")
        for dtype, count in sorted(s["by_type"].items()):
            table.add_row(dtype, str(count))
        console.print(table)

    if s.get("sources"):
        console.print(f"\n[bold]Sources ({len(s['sources'])}):[/bold]")
        for src in s["sources"]:
            console.print(f"  • {src}")


# ─── DELETE ───

@app.command()
def delete(
    source_file: str = typer.Argument(..., help="Source file name to remove from KB"),
    confirm: bool = typer.Option(False, "--yes", "-y", help="Skip confirmation"),
):
    """Remove all chunks from a source file."""
    if not confirm:
        typer.confirm(f"Delete all chunks from '{source_file}'?", abort=True)

    store = _get_store()
    count = store.delete_by_source(source_file)
    console.print(f"[green]Deleted {count} chunks from '{source_file}'[/green]")

    with console.status("Rebuilding BM25 index..."):
        searcher = _get_searcher()
        searcher.build_bm25_index()


# ─── REBUILD INDEX ───

@app.command()
def rebuild():
    """Rebuild the BM25 search index."""
    searcher = _get_searcher()
    with console.status("Rebuilding BM25 index..."):
        searcher.build_bm25_index()
    console.print("[green]BM25 index rebuilt successfully.[/green]")


if __name__ == "__main__":
    app()
