import datetime as dt
import inspect
import os
import re
import sys
import time
import tkinter as Tk
from pathlib import Path
from textwrap import wrap

import matplotlib
import matplotlib.backends.tkagg as tkagg
import matplotlib.cm as cm
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import PySimpleGUI as sg
from matplotlib.backends.backend_tkagg import FigureCanvasAgg
import os.path
import sys

f_size = 12
now = dt.datetime.now().strftime("%Y-%m-%d_%I%M%p").upper()
plt.style.use("ggplot")

my_path = os.path.abspath(os.path.dirname(__file__))


def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    base_path = getattr(sys, '_MEIPASS', os.path.dirname(
        os.path.abspath(__file__)))
    return os.path.join(base_path, relative_path)


# WELCOME_PIC = resource_path('vt_welcome.png')
# TASK_LIST_PIC = resource_path('ppt_lst_csv_ex.png')
# RESP_LIST_PIC = resource_path('att_lst_csv_ex.png')
# VT_BG_PIC = resource_path('vt_background.png')
# ICON_PIC = resource_path('just_E.ico')

WELCOME_PIC = os.path.join(my_path, 'vt_welcome.png')
TASK_LIST_PIC = os.path.join(my_path, 'task_list.png')
RESP_LIST_PIC = os.path.join(my_path, 'respondent_list.png')
VT_BG_PIC = os.path.join(my_path, 'vt_background.png')
ICON_PIC = os.path.join(my_path, 'vt.ico')

ACCENT_COLOR = "#003399"


def Splash():
    '''
    Splash screen
    '''
    layout = [[sg.Image(WELCOME_PIC, pad=(0, 0))]]
    sg.Window(
        "Value Tool",
        auto_close_duration=3,
        auto_close=True,
        element_padding=(0, 0),
        size=(500, 360),
        no_titlebar=True,
    ).Layout(layout).Read()


def Directions_Task_Tbl():
    '''
    Instruction prompt for importing ask list
    '''
    layout = [
        [
            sg.Text(
                "First select a .xlsx file with program,\n" +
                "project and task information\n" +
                "in the following schema:\n" +
                "\n(see: test_task_list.xlsx for example)\n",
                text_color=ACCENT_COLOR,
                font=("current 16"),
            )
        ],
        [sg.Image(TASK_LIST_PIC)],
        [sg.T("")],
        [sg.OK(bind_return_key=True, size=(10, 1)),
         sg.T(""), sg.Exit(size=(10, 1))],
    ]
    window = sg.Window(
        "Load Task List",
        grab_anywhere=True,
        icon=ICON_PIC,
        no_titlebar=False
    ).Layout(layout)
    window.Finalize()

    while True:
        event, _ = window.Read()
        if event is "OK" or event is None:
            window.Close()
            break
        elif event is "Exit":
            return True


def get_xlsx():
    '''
    Importing xlsx function from file prompt
    '''
    filename = sg.PopupGetFile(
        '',
        initial_folder=os.path.dirname(os.path.abspath(__file__)),
        no_window=True,
        file_types=(("xlsx", "*.xlsx"),),
        icon=ICON_PIC,
    )

    while not filename:
        filename = get_xlsx()

    return filename


def Task_Tbl():
    sg.SetOptions(auto_size_buttons=True)
    filename = get_xlsx()
    # --- populate table with file contents --- #
    if filename is not None:
        try:
            task_df = pd.read_excel(filename)
            for name in ['prog_num', 'prog_name', 'proj_num', 'proj_name', 'task']:
                if name not in task_df.columns:
                    sg.PopupError(
                        "Make sure you've selected the correct table"
                    )
                    filename = get_xlsx()
                    task_df = pd.read_excel(filename)
        except Exception as e:
            sg.PopupError(
                "Error reading file:\n" +
                str(e) +
                "\nPlease check field names."
            )
            sys.exit(69)

    try:
        task_df.columns = task_df.columns.str.replace("[^a-zA-Z_]", "")
    except Exception as e:
        sg.Popup("Problem loading data")

    return task_df


def value_lookup(x):
    vt_val_dict = {
        (1, 1): 1,
        (1, 2): 1,
        (1, 3): 2,
        (1, 4): 3,
        (1, 5): 3,
        (2, 1): 1,
        (2, 2): 1,
        (2, 3): 2,
        (2, 4): 4,
        (2, 5): 4,
        (3, 1): 1,
        (3, 2): 2,
        (3, 3): 3,
        (3, 4): 5,
        (3, 5): 5,
        (4, 1): 2,
        (4, 2): 3,
        (4, 3): 4,
        (4, 4): 5,
        (4, 5): 5,
        (5, 1): 2,
        (5, 2): 4,
        (5, 3): 5,
        (5, 4): 5,
        (5, 5): 5,
    }
    return vt_val_dict[x]


def Directions_Respondent_Tbl():
    layout = [
        [
            sg.Text(
                "Now select a .xlsx file with respondent\n" +
                "information in the following schema:\n" +
                "\n(see: test_respondent_list.xlsx for example)\n",
                text_color=ACCENT_COLOR,
                font=("current 16"),
            )
        ],
        [sg.Image(RESP_LIST_PIC)],
        [sg.T("")],
        [sg.OK(bind_return_key=True, size=(10, 1)),
         sg.T(""), sg.Exit(size=(10, 1))],
    ]
    window = sg.Window(
        "Load Respondent List",
        grab_anywhere=True,
        icon=ICON_PIC,
    ).Layout(layout)
    window.Finalize()

    while True:
        event, _ = window.Read()
        if event is "OK":
            window.Close()
            break
        elif event is "Exit":
            return True


def Respondent_Tbl():
    sg.SetOptions(auto_size_buttons=True)
    filename = get_xlsx()
    # --- populate table with file contents --- #
    if filename is not None:
        try:
            att_df = pd.read_excel(filename)
            for name in ['company', 'email', 'prog_num', 'prog_name', 'proj_num', 'proj_name', 'task', 'usability', 'roi', 'relevance', 'likelihood']:
                if name not in att_df.columns:
                    sg.PopupError(
                        "Make sure you've selected the correct table"
                    )
                    filename = get_xlsx()
                    att_df = pd.read_excel(filename)
        except Exception as e:
            sg.PopupError("Error reading file:" + str(e))
            sys.exit(69)

    try:
        att_df.columns = att_df.columns.str.replace("[^a-zA-Z_]", "")
    except Exception as e:
        sg.Popup("Problem loading data")

    att_df.dropna(inplace=True)
    att_df["usability"] = att_df["usability"].astype("int")
    att_df["roi"] = att_df["roi"].astype("int")
    att_df["relevance"] = att_df["relevance"].astype("int")
    att_df["likelihood"] = att_df["likelihood"].astype("int")

    val_list = []
    for _, row in att_df.iterrows():
        val_list.append(value_lookup((row["roi"], row["relevance"])))

    att_df["value"] = val_list

    att_df = att_df.groupby(["task", "company", "proj_num", "proj_name"]).agg(
        {
            "usability": "mean",
            "roi": "mean",
            "relevance": "mean",
            "likelihood": "mean",
            "company": "count",
            "value": 'mean'
        }
    )

    att_df.rename(index=str, columns={
                  "company": "respondent_count"}, inplace=True)
    att_df.reset_index(inplace=True)
    att_df = att_df.round(2)

    return att_df


def Grouping(group_col_list, task_tbl, respondent_tbl):
    try:
        merged = pd.merge(
            task_tbl,
            respondent_tbl,
            how="inner",
            on=["proj_name", "task"],
            suffixes=("_", ""),
            sort=False,
        )
        mask = ~merged.columns.str.endswith("_")
        merged = merged[merged.columns[mask]]
        groups = merged.groupby(group_col_list)

    except Exception as e:
        sg.Popup(
            'Please restart software\n' +
            'Error in merging of tables:\n' +
            'Failed in "Grouping" function: ' +
            str(e)
        )

    return groups


def Chrt_Menu(groups):
    def Bbl_Chrt(name, group):
        img = plt.imread(VT_BG_PIC)
        fig, ax = plt.subplots(figsize=(8, 8))  # inches
        ax.imshow(img, extent=[0, 5.0, 0, 5.0])
        plt.gcf()
        plt.xticks(fontsize=f_size)
        plt.yticks(fontsize=f_size)
        plt.axis((0.0, 5.0, 0, 5.0))

        N = len(group)

        # need qualitative colormap
        cmap = cm.get_cmap('nipy_spectral')
        colors = cmap(np.linspace(0, 1, N))

        # Use those colors as the color argument
        ax.scatter(
            group["usability"],
            group["value"],
            s=group["likelihood"] * 2000,
            color=colors,
            label=group["task"],
            alpha=0.5,
        )

        plt.xlabel("Average Usability", fontsize=f_size)
        plt.ylabel("Average Value", fontsize=f_size)
        if isinstance(name, tuple):
            fig.suptitle(name[0] + "\n" +
                         "\n".join(wrap(name[1])), fontsize=f_size)
        else:
            fig.suptitle("\n".join(wrap(name)), fontsize=f_size)

        return fig, colors

    fig_dict = {}
    for name, group in groups:
        pretty_name = str(name[0]) + ": " + str(name[1].replace("{", ""))
        fig_dict[pretty_name] = Bbl_Chrt, pretty_name, group

    # The magic function that makes it possible....
    # glues together tkinter and pyplot using Canvas Widget
    def draw_figure(canvas, figure, loc=(0, 0)):
        """
        Draw a matplotlib figure onto a Tk canvas
        loc: location of top-left corner of figure on canvas in pixels.
        Inspired by matplotlib source: lib/matplotlib/backends/backend_tkagg.py
        """
        figure_canvas_agg = FigureCanvasAgg(figure)
        figure_canvas_agg.draw()
        # figure_x, figure_y - unused but they used to be in front along w/..
        _, _, figure_w, figure_h = figure.bbox.bounds
        figure_w, figure_h = int(figure_w), int(figure_h)
        photo = Tk.PhotoImage(master=canvas, width=figure_w, height=figure_h)

        # Position: convert from top-left anchor to center anchor
        canvas.create_image(loc[0] + figure_w / 2,
                            loc[1] + figure_h / 2, image=photo)

        # Unfortunately, there's no accessor for the pointer
        # to the native renderer
        tkagg.blit(photo,
                   figure_canvas_agg.get_renderer()._renderer,
                   colormode=2)

        # Return a handle which contains a reference to the photo object
        # which must be kept live or else the picture disappears
        return photo

    # -------------------------------- GUI Starts Here -----------------------#
    # fig = your figure you want to display.  Assumption is that 'fig' holds  #
    # the information to display.                                             #

    figure_w, figure_h = 800, 800  # pixels
    # define the form layout
    listbox_values = [key for key in fig_dict.keys()]

    col_listbox = [
        [
            sg.Listbox(
                values=listbox_values,
                change_submits=True,
                bind_return_key=True,
                size=(125, 20),
                key="grouped_data",
            )
        ],
        [sg.Text("Current Plot Table", font=(
            "current 20"), text_color=ACCENT_COLOR)],
        [
            sg.Table(
                values=[["XX.XXX", "Company", "1", "5", "5", "5", "5", "5"]],
                headings=[
                    "Project Number",
                    "Company",
                    "Respondent Count",
                    "Average Usability",
                    "Average ROI",
                    "Average Relevance",
                    "Average Likelihood",
                    "Average Value",
                ],
                num_rows=15,
                auto_size_columns=False,
                display_row_numbers=False,
                justification="center",
                key="Plot_Tbl",
            )
        ],
        [sg.Text("Select Task Above Before Saving",
                 font=("current 20"),
                 text_color=ACCENT_COLOR)],
        [
            sg.Save(
                "Save Current Table",
                key="save_tbl",
                size=(20, 1),
                font=("current 16")
            ),
            sg.T(""),
            sg.Save(
                "Save Current Figure",
                key="save_fig",
                size=(20, 1),
                font=("current 16")
            ),
            sg.T(""),
            sg.Save(
                "Save Both",
                key="save_both",
                size=(10, 1),
                font=("current 16")
            ),
            sg.T(""),
            sg.Save(
                "Save All",
                key="save_all",
                size=(10, 1),
                font=("current 16")
            ),
            sg.T(""),
        ],
        [
            sg.Button(
                "Start Over",
                size=(10, 1),
                font=("current 16"),
                key="start_over"
            ),
            sg.T(""),
            sg.Exit(size=(10, 1),
                    font=("current 16")), ]
    ]

    col_canvas = sg.Column(
        [[sg.Canvas(
            size=(figure_w, figure_h),
            key="canvas",
        )]])

    layout = [
        [
            sg.Text(
                " Group List (select item to display)",
                font=("current 20"),
                text_color=ACCENT_COLOR,
            )
        ],
        [sg.Column(col_listbox), col_canvas],
    ]

    # create the form and show it without the plot
    window = sg.Window(
        title="Value Tool Plot",
        resizable=True,
        grab_anywhere=False,
        no_titlebar=False,
        icon=ICON_PIC,
        location=(100, 100),
    ).Layout(layout)
    window.Finalize()

    canvas_elem = window.FindElement("canvas")

    while True:
        event, values = window.Read()
        plt.close()

        try:
            choice = values["grouped_data"][0]
            func, name, group = fig_dict[choice]
            window.FindElement("Plot_Tbl").Update(
                values=group[
                    [
                        "proj_num",
                        "company",
                        "respondent_count",
                        "usability",
                        "roi",
                        "relevance",
                        "likelihood",
                        "value",
                    ]
                ].values.tolist()
            )
        except:
            pass

        # show it all again and get buttons
        plt.clf()
        fig, colors = func(name, group)
        _ = draw_figure(canvas_elem.TKCanvas, fig)

        group_list = group.company.values.tolist()
        group_color_list = []
        for grp_name, col in zip(group_list, colors):
            col_mapped = col * 255
            col_mapped = col_mapped.round(0)
            col_mapped = [int(i) for i in col_mapped]
            col_mapped_hex = '#%02x%02x%02x' % (col_mapped[0],
                                                col_mapped[1],
                                                col_mapped[2])

            group_color_list.append([grp_name, col_mapped[:3], col_mapped_hex])

        group_color_df = pd.DataFrame(
            group_color_list, columns=["company",
                                       "bubble_color_rgb",
                                       "bubble_color_hex"]
        )

        def save_current_fig(name, fig):
            try:
                save_fig_name = (name.replace("[^a-zA-Z]", "_")
                                 .replace("(", "")
                                 .replace(")", "")
                                 .replace(": ", "_")
                                 .replace("-", "")
                                 .replace(" / ", "_")
                                 .replace(" ", "_")
                                 .replace(":", "_")
                                 .replace(".", "_")
                                 .lower()[:30] + '.png')
                fig.savefig(save_fig_name)
            except Exception:
                pass

        def save_current_tbl(name, group):
            try:
                group_list = group.company.values.tolist()
                group_color_list = []
                for grp_name, col in zip(group_list, colors):
                    col_mapped = col * 255
                    col_mapped = col_mapped.round(0)
                    col_mapped = [int(i) for i in col_mapped]
                    col_mapped_hex = '#%02x%02x%02x' % (col_mapped[0],
                                                        col_mapped[1],
                                                        col_mapped[2])

                    group_color_list.append([grp_name,
                                             col_mapped[:3],
                                             col_mapped_hex])

                group_color_df = pd.DataFrame(
                    group_color_list, columns=["company",
                                               "bubble_color_rgb",
                                               "bubble_color_hex"]
                )

                save_tbl_name = (name.replace("[^a-zA-Z]", "_")
                                 .replace('&', '')
                                 .replace("(", "")
                                 .replace(")", "")
                                 .replace(": ", "_")
                                 .replace("-", "")
                                 .replace(" / ", "_")
                                 .replace(" ", "_")
                                 .replace(":", "_")
                                 .replace(".", "_")
                                 .lower()[:30] + '.xlsx')

                all_plot_data_df = pd.merge(
                    group, group_color_df, how="left", on="company")
                all_plot_data_df = all_plot_data_df.rename(
                    index=str,
                    columns={
                        "prog_num": "Program Number",
                        "prog_name": "Program Name",
                        "proj_num": "Project Number",
                        "proj_name": "Project Name",
                        "task": "Task Name",
                        "company": "Company",
                        "respondent_count": "Respondent Count",
                        "usability": "Average Usability",
                        "roi": "Average ROI",
                        "relevance": "Average Relevance",
                        "likelihood": "Average Likelihood",
                        "value": "Average Value",
                        "bubble_color_rgb": "Bubble Color [R, G, B]",
                        "bubble_color_hex": "Bubble Color Hex",
                    },
                )

                all_plot_data_df = all_plot_data_df[
                    [
                        "Program Number",
                        "Program Name",
                        "Project Number",
                        "Project Name",
                        "Task Name",
                        "Company",
                        "Respondent Count",
                        "Average Usability",
                        "Average ROI",
                        "Average Relevance",
                        "Average Likelihood",
                        "Average Value",
                        "Bubble Color [R, G, B]",
                        "Bubble Color Hex"
                    ]
                ]

                writer = pd.ExcelWriter(save_tbl_name,
                                        engine='xlsxwriter')
                all_plot_data_df.to_excel(writer,
                                          sheet_name=save_tbl_name[:30],
                                          index=False)
                writer.save()

            except Exception as e:
                print(str(e))

        def save_all_figs_and_tbls():
            i = 0
            for name, group in groups:
                pretty_name = (name[0].replace(".", "_").lower() + '_' +
                               name[1].replace("[^a-zA-Z]", "_")
                               .replace('&', '')
                               .replace("(", "")
                               .replace("(", "")
                               .replace(")", "")
                               .replace(": ", "_")
                               .replace("-", "")
                               .replace(" / ", "_")
                               .replace(" ", "_")
                               .replace(":", "_")
                               .replace(".", "_").lower()[:22])
                Bbl_Chrt(name, group)
                plt.savefig(pretty_name + '.png')
                save_current_tbl(name[0] + '_' + name[1], group)
                plt.close()
                print('Saving: ' + name[0] + '_' + name[1])

                sg.OneLineProgressMeter('Saving Plots and Tables...',
                                        current_value=i+1,
                                        max_value=len(groups),
                                        key='progress')
                i += 1

        if event is None or event is "Exit":
            break
        elif event is "save_tbl":
            save_current_tbl(name, group)
        elif event is "save_fig":
            save_current_fig(name, fig)
        elif event is 'save_both':
            save_current_fig(name, fig)
            save_current_tbl(name, group)
        elif event is 'save_all':
            save_all_figs_and_tbls()
        elif event is "start_over":
            window.Close()
            return True

    window.Close()


if __name__ == "__main__":
    start_over = True
    direct_ex = False
    while start_over:
        sg.SetOptions(
            icon=ICON_PIC,
            background_color="white",
            element_padding=(0, 0),
            text_element_background_color="white",
            element_background_color="white",
            button_color=(ACCENT_COLOR, "white"),
            window_location=(700, 200),
        )
        Splash()
        direct_ex = Directions_Task_Tbl()
        if direct_ex:
            break
        task_tbl = Task_Tbl()
        direct_ex = Directions_Respondent_Tbl()
        if direct_ex:
            break
        respondent_tbl = Respondent_Tbl()
        group_col_list = ["proj_num", "task"]
        groups = Grouping(group_col_list, task_tbl, respondent_tbl)
        start_over = Chrt_Menu(groups)
