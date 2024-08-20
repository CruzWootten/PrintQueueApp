import win32print
import win32api
from textual.app import App
from textual.widgets import Header, Footer, ScrollView, Static, Button

class PrintQueueApp(App):
    async def on_load(self) -> None:
        await self.bind("q", "quit", "Quit")

    async def on_mount(self) -> None:
        await self.view.dock(Header(), edge="top")
        await self.view.dock(Footer(), edge="bottom")

        self.print_queue = []
        self.queue_view = ScrollView(Static("Print Queue is empty."))
        await self.view.dock(self.queue_view, edge="center")

        add_button = Button("Add to Queue")
        add_button.when_clicked = self.add_to_queue
        print_button = Button("Print Queue")
        print_button.when_clicked = self.print_queue_items
        await self.view.dock(add_button, edge="left")
        await self.view.dock(print_button, edge="right")

    def add_to_queue(self) -> None:
        document = "Work Order 123"
        self.print_queue.append(document)
        self.queue_view.content.set_text(f"Print Queue: {', '.join(self.print_queue)}")

    def print_queue_items(self) -> None:
        while self.print_queue:
            document = self.print_queue.pop(0)
            try:
                printer_name = win32print.GetDefaultPrinter()
                hPrinter = win32print.OpenPrinter(printer_name)
                try:
                    win32print.StartDocPrinter(hPrinter, 1, (document, None, "RAW"))
                    win32print.StartPagePrinter(hPrinter)
                    win32print.WritePrinter(hPrinter, document.encode('utf-8'))
                    win32print.EndPagePrinter(hPrinter)
                    win32print.EndDocPrinter(hPrinter)
                    self.queue_view.content.set_text(f"Print Queue: {', '.join(self.print_queue)}")
                finally:
                    win32print.ClosePrinter(hPrinter)
            except Exception as e:
                self.queue_view.content.set_text(f"Error printing {document}: {e}")

if __name__ == "__main__":
    PrintQueueApp.run()
