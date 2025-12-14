"""Main application - Voice-controlled Lark task agent."""

import os
import sys
from typing import Optional
from dotenv import load_dotenv
from rich.console import Console
from rich.panel import Panel
from rich.prompt import Prompt, Confirm
from rich.table import Table

from .lark_client import LarkClient
from .agent import TaskAgent
from .speech_recognizer import SimpleSpeechInput, create_recognizer


console = Console()


def load_config() -> dict:
    """Load configuration from environment variables."""
    load_dotenv()

    config = {
        "lark_app_id": os.getenv("LARK_APP_ID"),
        "lark_app_secret": os.getenv("LARK_APP_SECRET"),
        "anthropic_api_key": os.getenv("ANTHROPIC_API_KEY"),
        "default_tasklist_id": os.getenv("LARK_DEFAULT_TASKLIST_ID"),
    }

    missing = [k for k, v in config.items() if not v and k != "default_tasklist_id"]
    if missing:
        console.print(f"[red]Missing required environment variables: {', '.join(missing)}[/red]")
        console.print("Please copy .env.example to .env and fill in your credentials.")
        sys.exit(1)

    return config


def select_tasklist(lark: LarkClient) -> Optional[str]:
    """Let user select a task list."""
    try:
        tasklists = lark.get_tasklists()
    except Exception as e:
        console.print(f"[yellow]タスクリストの取得に失敗しました: {e}[/yellow]")
        return None

    if not tasklists:
        console.print("[yellow]タスクリストがありません。タスクはデフォルトリストに作成されます。[/yellow]")
        return None

    table = Table(title="タスクリスト")
    table.add_column("#", style="cyan")
    table.add_column("名前", style="green")

    for i, tl in enumerate(tasklists, 1):
        table.add_row(str(i), tl.get("name", "無題"))

    console.print(table)

    choice = Prompt.ask(
        "使用するタスクリストの番号を入力 (Enterでスキップ)",
        default=""
    )

    if choice.isdigit():
        idx = int(choice) - 1
        if 0 <= idx < len(tasklists):
            selected = tasklists[idx]
            console.print(f"[green]選択: {selected.get('name')}[/green]")
            return selected.get("guid")

    return None


def run_voice_mode(agent: TaskAgent, speech_input):
    """Run the agent in voice input mode."""
    console.print(Panel(
        "[bold cyan]🎤 音声入力モード[/bold cyan]\n\n"
        "• 話しかけてタスクを管理できます\n"
        "• 'やめる' または 'exit' で終了\n"
        "• macOS Dictation: Fnキーを2回押す",
        title="Voice Mode"
    ))

    while True:
        # Get input
        if hasattr(speech_input, 'listen_with_prompt'):
            user_input = speech_input.listen_with_prompt("話しかけてください...")
        elif hasattr(speech_input, 'listen_once'):
            console.print("\n[cyan]🎤 聞いています...[/cyan]")
            user_input = speech_input.listen_once(timeout=15.0)
        else:
            user_input = Prompt.ask("\n[cyan]入力[/cyan]")

        if not user_input:
            console.print("[yellow]入力がありませんでした[/yellow]")
            continue

        # Check for exit commands
        if user_input.lower() in ["exit", "quit", "やめる", "終了", "おわり"]:
            console.print("[green]終了します。[/green]")
            break

        console.print(f"[dim]入力: {user_input}[/dim]")

        # Process with agent
        with console.status("[bold green]処理中..."):
            result = agent.process_input(user_input)

        console.print(Panel(result, title="結果", border_style="green"))


def run_interactive_mode(agent: TaskAgent):
    """Run the agent in text input mode."""
    console.print(Panel(
        "[bold cyan]💬 テキスト入力モード[/bold cyan]\n\n"
        "• タスク管理コマンドをテキストで入力\n"
        "• 'exit' または 'quit' で終了\n"
        "• 'clear' で会話履歴をクリア",
        title="Interactive Mode"
    ))

    while True:
        try:
            user_input = Prompt.ask("\n[cyan]あなた[/cyan]")
        except (KeyboardInterrupt, EOFError):
            break

        if not user_input.strip():
            continue

        if user_input.lower() in ["exit", "quit"]:
            break

        if user_input.lower() == "clear":
            agent.clear_history()
            console.print("[green]会話履歴をクリアしました[/green]")
            continue

        # Process with agent
        with console.status("[bold green]処理中..."):
            result = agent.process_input(user_input)

        console.print(Panel(result, title="アシスタント", border_style="green"))


def main():
    """Main entry point."""
    console.print(Panel.fit(
        "[bold blue]Lark Voice Agent[/bold blue]\n"
        "音声でLarkタスクを管理するAIエージェント",
        border_style="blue"
    ))

    # Load configuration
    config = load_config()

    # Initialize Lark client
    console.print("[dim]Lark APIに接続中...[/dim]")
    lark = LarkClient(
        app_id=config["lark_app_id"],
        app_secret=config["lark_app_secret"]
    )

    # Select task list
    tasklist_guid = config.get("default_tasklist_id")
    if not tasklist_guid:
        tasklist_guid = select_tasklist(lark)

    # Initialize agent
    agent = TaskAgent(
        anthropic_api_key=config["anthropic_api_key"],
        lark_client=lark
    )
    if tasklist_guid:
        agent.set_default_tasklist(tasklist_guid)

    # Select input mode
    console.print("\n[bold]入力モードを選択:[/bold]")
    console.print("  1. 🎤 音声入力 (macOS Dictationを使用)")
    console.print("  2. 💬 テキスト入力")

    mode = Prompt.ask("モード", choices=["1", "2"], default="2")

    try:
        if mode == "1":
            # Initialize speech recognizer
            speech_input = SimpleSpeechInput()
            run_voice_mode(agent, speech_input)
        else:
            run_interactive_mode(agent)
    except KeyboardInterrupt:
        console.print("\n[yellow]中断されました[/yellow]")
    finally:
        lark.close()
        console.print("[green]接続を閉じました。[/green]")


if __name__ == "__main__":
    main()
