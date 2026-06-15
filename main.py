import argparse 
import os 
from rich.console import Console 
from rich.prompt import Prompt 
from agent import run_agent 
from state import SessionState 
from config import DEFAULT_WORKDIR 

console = Console()

def main():

    parser = argparse.ArgumentParser(description="OfficeBot - AI Office assistant")
    parser.add_argument(
        "--workdir",
        default=DEFAULT_WORKDIR,
        help="Directory to read/save from. Defaults to ~/officebot_files"
    )
    args = parser.parse_args()

    os.makedirs(args.workdir, exist_ok=True)

    SESSION_FILE = os.path.join(args.workdir, "session.json")

    state = SessionState(workdir=args.workdir)
    state.load_from_file(SESSION_FILE)


    console.print("\n [bold green]OfficeBot[/bold green] - Word & Excel AI Assistant")
    console.print(f"Working directory: [cyan]{args.workdir}[/cyan]")
    console.print("Type your instruction. Type [bold]exit[/bold] to quit. \n")

    while True:
        user_input = Prompt.ask("[bold]You[/bold]")

        if user_input.strip().lower() in ("exit", "quit"):
            state.save_to_file(SESSION_FILE)     
            console.print("Bye.")
            break 
        if not user_input.strip():
            continue 

        console.print("\n[bold yellow]Bot[/bold yellow]")

        response = run_agent(user_input, state=state)
        console.print(response)
        print()


if __name__ == "__main__":
    main()


