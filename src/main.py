import flet as ft
import os
import sys
from datetime import datetime as dt

from pdf import PDFToExcel

is_converting = False

def main(page: ft.Page):
    # Configure app (main page/the only page) initial properties
    page.title='PCMS PDF Table Extractor'
    page.window_width=400
    page.window_height=400
    page.window_resizable=False
    page.window_maximizable=False
    page.theme_mode = ft.ThemeMode.LIGHT
    page.window_center()
    page.scroll = 'AUTO'


    # Open Notification Dialog
    def open_dlg(msg: str):
        dlg = ft.AlertDialog(
            title=ft.Text(msg), 
            # on_dismiss=lambda e: print("Open file Dialog dismissed!")
        )
        page.dialog = dlg
        dlg.open = True 
        page.update()

    # Get PDF source file location
    def select_source_file_result(e: ft.FilePickerResultEvent):
        if e.files == None:
            print('None file selected!')
            return
        source_file.value = e.files[0].path

        convert_button.disabled = False if output_file.value != '' or is_converting else True

        source_file.update()
        convert_button.update()

    # Get Excel output file location
    def select_output_file_result(e: ft.FilePickerResultEvent):
        if e.path == None:
            print('None file selected!')
            return
        output_file.value = e.path

        convert_button.disabled = False if source_file.value != '' or is_converting else True

        output_file.update()
        convert_button.update()

    # Convert and save file! (Main function of the app)
    def convert(src_file: str, out_file: str):
        convert_button.disabled = True
        select_pdf_button.disabled = True
        select_out_button.disabled = True
        status_text.value = 'Status: Converting...'
        page.update()

        is_converting = True
        
        try:
            PDFToExcel(source_file.value, output_file.value)
        except Exception as e:
            open_dlg("Error occured! " + str(e))
        
        open_dlg("Convert completed!")

        status_text.value = 'Status: Completed. Idle now.'
        convert_button.disabled = False
        select_pdf_button.disabled = False
        select_out_button.disabled = False
        page.update()

        is_converting = False

        os.startfile(out_file)

    # Initialize app controls
    source_file = ft.TextField(read_only=True, label='PDF File source:')
    output_file = ft.TextField(read_only=True, label='Excel File output:')
    select_source_files_dialog = ft.FilePicker(on_result=select_source_file_result)
    select_output_file_dialog = ft.FilePicker(on_result=select_output_file_result)
    convert_button = ft.ElevatedButton("Convert and save file", on_click=lambda _: convert(source_file.value, output_file.value), disabled=True)
    status_text = ft.Text("Status: Idle")
    select_pdf_button = ft.IconButton(
                        icon=ft.icons.UPLOAD_FILE,
                        bgcolor=ft.colors.LIGHT_BLUE_100,
                        on_click=lambda _: select_source_files_dialog.pick_files(
                            allow_multiple=False,
                            allowed_extensions=['pdf']
                        ),
                    )
    select_out_button = ft.IconButton(
                        icon=ft.icons.UPLOAD_FILE,
                        bgcolor=ft.colors.LIGHT_BLUE_100,
                        on_click=lambda _: select_output_file_dialog.save_file(
                            file_name = f"extracted {dt.now().strftime('%Y-%m-%d_%H-%M-%S')}.xlsx",
                        ),
    )

    # add file picker control dialog (required to open a file dialog)
    page.overlay.append(select_source_files_dialog)
    page.overlay.append(select_output_file_dialog)


    page.add(
        ft.Row(
            [
                # ft.Container(ft.Text("Source file:"), col=3),
                ft.Container(source_file, padding=5, expand=5),
                ft.Container(
                    select_pdf_button,
                    padding=5,
                    expand=1
                )
            ],
        ),
        ft.Row(
            [
                # ft.Container(ft.Text("Source file:"), col=3),
                ft.Container(output_file, padding=5, expand=5),
                ft.Container(
                    select_out_button,
                    padding=5,
                    expand=1
                )
            ],   
        ),
        ft.Row(
            [
                # ft.Container(ft.Text("Source file:"), col=3),
                ft.Container(convert_button, padding=5, expand=1),
            ],   
        ),
        ft.Row(
            [
                ft.Container(status_text, padding=5, expand=1),
            ],
        )
    )



ft.app(target=main)