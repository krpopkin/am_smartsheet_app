import reflex as rx
from . import smartsheet_login
from . import download_program_plan
from . import create_wip_report
import sys
import asyncio
import os
from dotenv import load_dotenv


class State(rx.State):
    """The app state."""
    current_content: str = " "
    output_lines: list[str] = []
    is_running: bool = False

    def show_wip_report(self):
        """Trigger the WIP report background task."""
        return State.run_wip_report
    
    @rx.event(background=True)
    async def run_wip_report(self):
        """Run the WIP report as a background task."""
        async with self:
            self.current_content = "WIP report - Running..."
            self.output_lines = []
            self.is_running = True
        
        output_collector = []
        original_stdout = sys.stdout
        playwright_instance = None
        browser = None
        
        try:
            # Load .env file
            load_dotenv()
            
            sys.stdout = OutputCapture(original_stdout, output_collector)
            
            # Step 1: Login and get browser session (already async, no executor needed)
            p, browser, context, page = await smartsheet_login.main()
            playwright_instance = p
            
            async with self:
                self.output_lines = list(output_collector)
            
            # Step 2: Download program plan using the same browser session (already async)
            await download_program_plan.main(page)
            
            async with self:
                self.output_lines = list(output_collector)
            
            # Step 3: Create WIP report (sync function, use executor)
            loop = asyncio.get_event_loop()
            await loop.run_in_executor(None, create_wip_report.main)
            
            async with self:
                self.output_lines = list(output_collector)
            
            async with self:
                self.current_content = "WIP report - Completed!"
                self.is_running = False
            
        except Exception as e:
            async with self:
                self.output_lines = [*output_collector, f"Error: {str(e)}"]
                self.current_content = "WIP report - Error occurred"
                self.is_running = False
        finally:
            sys.stdout = original_stdout
            # Clean up browser resources (now async)
            if browser:
                await browser.close()
            if playwright_instance:
                await playwright_instance.stop()

    def show_update_ss(self):
        """Update content to Update SS."""
        self.current_content = "Update SS"
        self.output_lines = []
        self.is_running = False


class OutputCapture:
    """Custom class to capture print output to a list."""
    
    def __init__(self, original_stdout, collector):
        self.original_stdout = original_stdout
        self.collector = collector
    
    def write(self, text):
        """Write to both the original stdout and collect the output."""
        if text.strip():
            self.collector.append(text.rstrip())
        self.original_stdout.write(text)
        return len(text)
    
    def flush(self):
        """Flush the stream."""
        self.original_stdout.flush()


def index() -> rx.Component:
    """The main page."""
    return rx.container(
        rx.vstack(
            rx.heading("Smartsheet App", size="8"),
            rx.hstack(
                rx.button("WIP report", on_click=State.show_wip_report, size="3", disabled=State.is_running),
                rx.button("Update SS", on_click=State.show_update_ss, size="3", disabled=State.is_running),
                spacing="4",
                padding_y="4",
            ),
            rx.heading(State.current_content, size="6"),
            rx.cond(
                State.output_lines.length() > 0,
                rx.box(
                    rx.vstack(
                        rx.foreach(
                            State.output_lines,
                            lambda line: rx.text(line, font_family="monospace", font_size="14px", white_space="pre-wrap"),
                        ),
                        spacing="1",
                        align="start",
                        width="100%",
                    ),
                    width="100%",
                    max_height="600px",
                    overflow_y="auto",
                    padding="4",
                    border="1px solid #ccc",
                    border_radius="8px",
                    background_color="#f5f5f5",
                ),
            ),
            spacing="6",
            align="center",
            padding_y="8",
            width="100%",
        ),
        size="3",
    )


app = rx.App()
app.add_page(index)