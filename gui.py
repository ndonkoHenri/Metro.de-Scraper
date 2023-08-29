import flet as ft
from splinter import Browser
import utils
import time

src_path = ""
dst_path = ""

# todo: make the progress bar show the automation's progression in real time instead of it being inderteminate


def main(page: ft.Page):
    global src_path, dst_path
    
    # page settings/configurations
    page.title = "Metro Scraper"
    page.vertical_alignment = ft.CrossAxisAlignment.CENTER
    page.horizontal_alignment = ft.MainAxisAlignment.CENTER
    page.theme_mode = ft.ThemeMode.LIGHT
    page.splash = ft.ProgressBar(visible=False)
    page.window_height, page.window_width = 520, 488
    page.window_min_height, page.window_min_width = 420, 498
    page.spacing, page.padding = 20, 10
    page.scroll = ft.ScrollMode.HIDDEN
    page.window_center()

    def fp_result(e: ft.FilePickerResultEvent):
        global src_path, dst_path

        # pick files
        if e.files:
            src_text.value = e.files[0].name
            src_path = e.files[0].path
        elif e.path:
            a = e.path.split("\\")[-1]
            if ".xlsx" in a:
                dst_text.value = a
                dst_path = e.path
            else:
                dst_text.value = a + ".xlsx"
                dst_path = e.path + ".xlsx"

        # makes sure the user gave the source and destination paths; if yes enable the buttons for automation run
        if src_path and dst_path:
            row.disabled = False

        page.update()

    def start_automation(e: ft.ControlEvent):
        """
        The start_automation function is called when the user clicks on the 'Run Automation *' buttons at the top.
        It takes in a browser name as an argument and starts automation using that browser.


        Args:
            e: contains useful data on the control that triggered this function

        Returns:
            The logs and the browser object
        """
        page.splash.visible = True
        row.disabled = True
        page.update()
        page.show_snack_bar(
            ft.SnackBar(
                ft.Text(f"Running automation on your {e.control.data} browser..."),
                open=True,
                duration=10000,
                action="OK"
            )
        )

        b = Browser(e.control.data)

        col.controls.append(ft.Text(f"Starting Automation at {time.strftime('%H:%M:%S')}"))
        page.update()

        logs = utils.start_automation(b, src_path, dst_path)

        # show logs on UI
        if logs:
            for j in logs:
                col.controls.append(ft.Text(j))
                page.update()

        col.controls.append(
            ft.Text(
                f"- {time.strftime('%H:%M %p')} | Automation completed. Check your Result file for the results."
            )
        )

        page.show_snack_bar(
            ft.SnackBar(
                ft.Text(
                    f"- {time.strftime('%H:%M %p')} | Automation completed. Check Result file for the results."),
                open=True,
                duration=10000,
                action="OK"
            )
        )

        # make the window visible by bringing it to the front, to let the user know the execution is done
        page.window_to_front()
        page.splash.visible = False
        row.disabled = False
        page.update()

    def copy_all_logs(e):
        """
        Copies all the logs in the log area to clipboard.
        Then shows a snackbar to notify the user that logs have been successfully copied.
        """

        x = ""
        for i in col.controls:
            x += f"{i.value}\n"

        print(x)
        page.set_clipboard(x)
        page.show_snack_bar(ft.SnackBar(ft.Text(f"Copied logs to clipboard!"), duration=10000, action="OK"))

    def delete_all_logs(e):
        """
        Called when the user clicks on the 'Delete All Logs' button.
        It clears all the controls in col, which is a collection of all the logs that have been created till then.
        Then shows a snackbar to notify them that the logs were deleted successfully.
        """
        col.controls.clear()
        page.update()
        page.show_snack_bar(ft.SnackBar(ft.Text(f"All logs were deleted!"), duration=10000, action="OK"))

    fp = ft.FilePicker(on_result=fp_result)
    page.overlay.append(fp)

    page.add(
        ft.Row(
            controls=[
                ft.Container(
                    src_text := ft.Text("Select SOURCE file here", weight=ft.FontWeight.BOLD),
                    bgcolor=ft.colors.GREY_300,
                    height=60,
                    width=255,
                    alignment=ft.alignment.center,
                    on_click=lambda e: fp.pick_files(
                        "Source...",
                        file_type=ft.FilePickerFileType.CUSTOM,
                        allowed_extensions=["xlsx"]
                    )
                ),
                ft.FloatingActionButton(
                    content=ft.Icon(ft.icons.FILE_OPEN, color=ft.colors.YELLOW_600),
                    on_click=lambda e: fp.pick_files(
                        "Source...",
                        file_type=ft.FilePickerFileType.CUSTOM,
                        allowed_extensions=["xlsx"]
                    ),
                    bgcolor=ft.colors.GREEN_300,
                    tooltip="dok einreichen",
                    mini=True
                )
            ],
            alignment=ft.MainAxisAlignment.CENTER
        ),
        ft.Row(
            controls=[
                ft.Container(
                    dst_text := ft.Text("Select RESULT file destination here", weight=ft.FontWeight.BOLD),
                    bgcolor=ft.colors.GREY_300,
                    height=60,
                    width=255,
                    alignment=ft.alignment.center,
                    on_click=lambda e: fp.save_file(
                        "Select result file...",
                        file_name="Result.xlsx",
                        file_type=ft.FilePickerFileType.CUSTOM,
                        allowed_extensions=["xlsx"]
                    )
                ),
                ft.FloatingActionButton(
                    content=ft.Icon(ft.icons.FILE_OPEN, color=ft.colors.YELLOW_600),
                    on_click=lambda e: fp.save_file(
                        "Select result file...",
                        file_name="Result.xlsx",
                        file_type=ft.FilePickerFileType.CUSTOM,
                        allowed_extensions=["xlsx"]
                    ),
                    bgcolor=ft.colors.GREEN_300,
                    tooltip="load document",
                    mini=True
                )
            ],
            alignment=ft.MainAxisAlignment.CENTER
        ),
        row := ft.Row(
            controls=[
                ft.OutlinedButton("Run Automation on Edge", on_click=start_automation, data="edge"),
                ft.OutlinedButton("Run Automation on Chrome", on_click=start_automation, data="chrome")
            ],
            disabled=True,
            alignment=ft.MainAxisAlignment.SPACE_EVENLY
        ),
        ft.Text(
            size=20,
            weight=ft.FontWeight.BOLD,
            spans=[
                ft.TextSpan("Logs", style=ft.TextStyle(decoration=ft.TextDecoration.UNDERLINE))
            ]
        ),
        ft.Container(
            col := ft.Column(
                controls=[
                    ft.Text(
                        "Welcome! Select your source excel file (with 2 main columns: 'Metro Artikelnummer' and 'Links'), then run the automation by clicking the buttons above for your preferred browser (either Edge or Chrome)"
                    ),
                ],
                horizontal_alignment=ft.CrossAxisAlignment.STRETCH,
                auto_scroll=True,
                scroll="hidden"
            ),
            alignment=ft.alignment.top_left,
        ),
        ft.Row(
            controls=[
                ft.ElevatedButton(
                    "Copy Logs",
                    icon=ft.icons.COPY_ROUNDED,
                    on_click=copy_all_logs,
                    bgcolor=ft.colors.LIGHT_GREEN_ACCENT_700,
                    color=ft.colors.GREY_100
                ),
                ft.ElevatedButton(
                    "Delete Logs",
                    icon=ft.icons.DELETE_FOREVER,
                    on_click=delete_all_logs,
                    bgcolor=ft.colors.RED_ACCENT_700,
                    color=ft.colors.LIME_50
                )
            ],
            alignment=ft.MainAxisAlignment.START
        )
    )


ft.app(main)
