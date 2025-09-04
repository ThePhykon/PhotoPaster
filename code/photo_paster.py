import time
import argparse
import os
import win32clipboard
import win32con
import struct
import rich.console
import rich.progress
from rich.table import Table
from rich.prompt import Confirm
from rich.panel import Panel
from rich.text import Text
from rich.layout import Layout
from rich.live import Live
import sys
import math
import winsound

BATCH_DELAY = 15
BATCH_SIZE = 16  # Globale Variable für Batch-Größe

image_extensions = (".jpg", ".jpeg", ".png", ".bmp", ".gif", ".tiff", ".webp")
console = rich.console.Console()

def get_clipboard_file_paths():
    try:
        win32clipboard.OpenClipboard()
        if win32clipboard.IsClipboardFormatAvailable(win32con.CF_HDROP):
            files = win32clipboard.GetClipboardData(win32con.CF_HDROP)
            return list(files)
        else:
            print("Keine Dateipfade im Clipboard.")
            return []
    finally:
        win32clipboard.CloseClipboard()

def copy_files_to_clipboard(filepaths):
    # Setze fWide=1 für Unicode
    DROPFILES_STRUCT = struct.pack('IiiII', 20, 0, 0, 0, 1)
    
    # Erstelle die Liste der Dateipfade, null-separiert und doppeltes Nullbyte am Ende
    files = '\0'.join(filepaths) + '\0\0'
    files_bytes = files.encode('utf-16le')
    data = DROPFILES_STRUCT + files_bytes
    win32clipboard.OpenClipboard()
    win32clipboard.EmptyClipboard()
    win32clipboard.SetClipboardData(win32con.CF_HDROP, data)
    win32clipboard.CloseClipboard()

def create_counter_panel(current_batch, total_batches, batch_size, current_photos, total_photos):
    counter_text = Text()
    counter_text.append("Batch Progress: ", style="bold white")
    counter_text.append(f"{current_batch}/{total_batches}", style="bold cyan")
    counter_text.append(" | Photos: ", style="bold white")
    counter_text.append(f"{current_photos}/{total_photos}", style="bold green")
    counter_text.append(f" | Current Batch Size: ", style="bold white")
    counter_text.append(f"{batch_size}", style="bold yellow")
    
    return Panel(
        counter_text,
        title="[bold blue]Upload Status[/bold blue]",
        border_style="blue",
        padding=(0, 1)
    )

if __name__ == "__main__":
    argparser = argparse.ArgumentParser(description="Upload photos to WhatsApp Web.")
    argparser.add_argument("photos_folder", nargs="+", help="List of photo file paths to upload.")
    
    args = argparser.parse_args()
    photos_folder = args.photos_folder
    
    with console.status("[bold cyan]Gathering photos...") as status:
        photos = []
        for folder in photos_folder:
            files = os.listdir(folder)
            
            photos.extend([
                os.path.abspath(os.path.join(folder, photo))
                for photo in files
                if photo.lower().endswith(image_extensions)
            ])
    
    table = Table(title="Gefundene Fotos")
    table.add_column("Nr.", style="cyan", width=4)
    table.add_column("Dateipfad", style="magenta")
    for idx, photo in enumerate(photos, 1):
        table.add_row(str(idx), photo)
    console.print(table)
    
    console.print(f"[bold white]> [/bold white][green]Found [grey]{len(photos)} [white]photos in provided folders!")
    start = Confirm.ask("[bold green]Start upload process?")
    
    if not start:
        sys.exit(0)
    
    # Berechne die Anzahl der Batches mit globaler BATCH_SIZE
    total_batches = math.ceil(len(photos) / BATCH_SIZE)
    total_photos = len(photos)
    photos_processed = 0
    
    # Erstelle Layout
    layout = Layout()
    
    # *** WICHTIGE ÄNDERUNG HIER: Layout wird einmalig mit 3 Bereichen definiert ***
    layout.split_column(
        Layout(name="counter_area", size=3),
        Layout(name="message_area", size=3), # Neuer Bereich für Nachrichten/Countdown
        Layout(name="progress_area", size=3)
    )
    
    # Progress Bar Setup
    progress = rich.progress.Progress(
        rich.progress.SpinnerColumn(),
        rich.progress.TextColumn("[progress.description]{task.description}"),
        rich.progress.BarColumn(),
        rich.progress.TaskProgressColumn(),
        rich.progress.TimeElapsedColumn(),
        rich.progress.TimeRemainingColumn(),
    )
    
    task = progress.add_task("Processing batches...", total=total_batches)
    
    with Live(layout, refresh_per_second=8, console=console):
        # Initiales Setup der Inhalte
        layout["counter_area"].update(create_counter_panel(0, total_batches, 0, 0, total_photos))
        layout["progress_area"].update(progress)
        
        # Anfangs ist der Nachrichtenbereich "bereit" oder leer
        layout["message_area"].update(
            Panel(
                Text("Starting...", style="bold green"),
                title="[bold green]Status[/bold green]",
                border_style="green",
                padding=(0, 1)
            )
        )

        time.sleep(3)
        
        for batch_num in range(total_batches):
            # Sound-Benachrichtigung beim Start jedes Batches
            winsound.Beep(800, 300)  # 800Hz, 300ms
            
            # Aktueller Batch mit globaler BATCH_SIZE
            current_batch_num = batch_num + 1
            batch = photos[:BATCH_SIZE]
            photos = photos[BATCH_SIZE:]
            current_batch_size = len(batch)
            photos_processed += current_batch_size
            
            # Update Counter Panel
            counter_panel = create_counter_panel(
                current_batch_num, 
                total_batches, 
                current_batch_size, 
                photos_processed, 
                total_photos
            )
            layout["counter_area"].update(counter_panel) # Update den Inhalt des Zähler-Bereichs
            
            # Kopiere Batch ins Clipboard
            copy_files_to_clipboard(batch)
            
            # Countdown mit Live Update
            for i in range(BATCH_DELAY, 0, -1):
                countdown_text = Text()
                countdown_text.append("Batch copied! Next batch in ", style="bold white")
                countdown_text.append(f"{i}", style="bold red")
                countdown_text.append(" seconds...", style="bold white")
                
                countdown_panel = Panel(
                    countdown_text, 
                    title="[bold yellow]Waiting[/bold yellow]", 
                    border_style="yellow", 
                    padding=(0, 1)
                )
                
                # *** WICHTIGE ÄNDERUNG HIER: Aktualisiere nur den Inhalt des message_area ***
                layout["message_area"].update(countdown_panel)
                
                time.sleep(1)
            
            # Update Progress
            progress.update(task, advance=1)
            
            # *** WICHTIGE ÄNDERUNG HIER: Setze den message_area wieder auf "bereit" oder leer ***
            if batch_num < total_batches -1: # Nicht nach dem letzten Batch
                layout["message_area"].update(
                    Panel(
                        Text("Ready for next batch...", style="bold green"),
                        title="[bold green]Status[/bold green]",
                        border_style="green",
                        padding=(0, 1)
                    )
                )
            else: # Nach dem letzten Batch
                layout["message_area"].update(
                    Panel(
                        Text("All batches processed.", style="bold green"),
                        title="[bold green]Status[/bold green]",
                        border_style="green",
                        padding=(0, 1)
                    )
                )

    console.print("[bold green]✓ All batches processed successfully![/bold green]")