"""This is a GUI of xls2gd, which is a tool to convert Excel files to GDScript files."""

#! /usr/bin/env python
# -*- coding: utf-8 -*
# description: xls2gd GUI
# @copyright Hidden Moss, https://hiddenmoss.com/
# @author Yuancheng Zhang, https://github.com/endaye
# @see repo: https://github.com/hidden-moss/xls2gd

import sys
import os
import wx
import wx.richtext as rt
import wx.lib.agw.hyperlink as hl
import tool_xls2gd as x2l


__authors__ = ["Yuancheng Zhang"]
__copyright__ = "©2025 Hidden Moss"
__credits__ = ["Yuancheng Zhang"]
__license__ = "MIT"
__version__ = "v1.2.4"
__maintainer__ = "Yuancheng Zhang"
__status__ = "Production"


class MainFrame(wx.Frame):
    """Main frame of the GUI."""
    prefix_color = {
        "info": "blue",
        "error": "red",
        "sucess": "sea green",
        "failed": "red",
    }

    def __init__(self, parent, title):
        wx.Frame.__init__(self, parent, -1, title, size=wx.Size(720, 320))
        self.SetIcon(wx.Icon(self.resource_path("img\\icon.ico")))
        self.panel = wx.Panel(self)

        self.sizer_v = wx.BoxSizer(wx.VERTICAL)

        font1 = wx.Font(8, wx.MODERN, wx.NORMAL, wx.NORMAL, False, "Consolas")

        # 采用多行显示
        self.logs = rt.RichTextCtrl(self.panel, style=wx.TE_MULTILINE | wx.TE_READONLY)
        self.logs.SetFont(font1)

        # 1为响应容器改变大小，expand占据窗口的整个宽度
        self.sizer_v.Add(self.logs, 1, wx.ALIGN_TOP | wx.EXPAND)
        self.panel.SetSizerAndFit(self.sizer_v)

        # Config - input path
        self.sizer_cfg_1 = wx.BoxSizer(wx.HORIZONTAL)
        self.st1 = wx.StaticText(self.panel, label=" Input Path:", size=(130, -1))
        self.sizer_cfg_1.Add(self.st1, flag=wx.RIGHT, border=4)
        self.tc1 = wx.TextCtrl(self.panel)
        self.sizer_cfg_1.Add(self.tc1, proportion=1)

        # Config - output gdscript path
        self.sizer_cfg_2 = wx.BoxSizer(wx.HORIZONTAL)
        self.st2 = wx.StaticText(self.panel, label=" Output Path (*.gd):", size=(130, -1))
        self.sizer_cfg_2.Add(self.st2, flag=wx.RIGHT, border=4)
        self.tc2 = wx.TextCtrl(self.panel)
        self.sizer_cfg_2.Add(self.tc2, proportion=1)

        # Config - output gdscript name template
        self.sizer_cfg_3 = wx.BoxSizer(wx.HORIZONTAL)
        self.st3 = wx.StaticText(self.panel, label=" Name Template (*.gd):", size=(130, -1))
        self.sizer_cfg_3.Add(self.st3, flag=wx.RIGHT, border=4)
        self.tc3 = wx.TextCtrl(self.panel)
        self.sizer_cfg_3.Add(self.tc3, proportion=1)
        
        # Config - output gdscript name template
        self.sizer_cfg_4 = wx.BoxSizer(wx.HORIZONTAL)
        self.st4 = wx.StaticText(self.panel, label=" Output Path (*.csv):", size=(130, -1))
        self.sizer_cfg_4.Add(self.st4, flag=wx.RIGHT, border=4)
        self.tc4 = wx.TextCtrl(self.panel)
        self.sizer_cfg_4.Add(self.tc4, proportion=1)
        
        # Config - output gdscript name template
        self.sizer_cfg_5 = wx.BoxSizer(wx.HORIZONTAL)
        self.st5 = wx.StaticText(self.panel, label=" Name Template (*.csv):", size=(130, -1))
        self.sizer_cfg_5.Add(self.st5, flag=wx.RIGHT, border=4)
        self.tc5 = wx.TextCtrl(self.panel)
        self.sizer_cfg_5.Add(self.tc5, proportion=1)

        # add Config to parent panel
        self.sizer_v.Add(
            self.sizer_cfg_1, flag=wx.EXPAND | wx.LEFT | wx.RIGHT | wx.TOP, border=4
        )
        self.sizer_v.Add(
            self.sizer_cfg_2, flag=wx.EXPAND | wx.LEFT | wx.RIGHT | wx.TOP, border=4
        )
        self.sizer_v.Add(
            self.sizer_cfg_3, flag=wx.EXPAND | wx.LEFT | wx.RIGHT | wx.TOP, border=4
        )
        self.sizer_v.Add(
            self.sizer_cfg_4, flag=wx.EXPAND | wx.LEFT | wx.RIGHT | wx.TOP, border=4
        )
        self.sizer_v.Add(
            self.sizer_cfg_5, flag=wx.EXPAND | wx.LEFT | wx.RIGHT | wx.TOP, border=4
        )

        # Bottom sizer
        self.sizer_btm = wx.BoxSizer(wx.HORIZONTAL)
        self.sizer_btm_l = wx.BoxSizer(wx.VERTICAL)
        self.sizer_btm_r = wx.BoxSizer(wx.VERTICAL)
        self.sizer_btm_r_h = wx.BoxSizer(wx.HORIZONTAL)

        # Convert
        self.btn_convert = wx.Button(self.panel, label="Convert")
        self.Bind(wx.EVT_BUTTON, self.on_convert_click, self.btn_convert)
        self.sizer_btm_l.Add((-1, 4))
        self.sizer_btm_l.Add(self.btn_convert, flag=wx.LEFT, border=4)

        # Version
        version_str = __version__ + " "
        self.st_version = wx.StaticText(
            self.panel, label=version_str, size=(100, -1), style=wx.ALIGN_RIGHT
        )
        self.sizer_btm_r_h.Add(self.st_version, flag=wx.RIGHT, border=4)

        # Config
        self.cb_config = wx.CheckBox(self.panel, label="Config")
        self.Bind(wx.EVT_CHECKBOX, self.on_config_checked, self.cb_config)
        self.sizer_btm_r_h.Add(self.cb_config, flag=wx.RIGHT, border=4)

        # Copyright
        copyright_link_str = __copyright__
        self.st_copyright_link = hl.HyperLinkCtrl(
            self.panel,
            label=copyright_link_str,
            URL="http://www.hiddenmoss.com/",
            size=(240, -1),
            style=wx.ALIGN_RIGHT,
        )

        self.st_copyright_link.EnableRollover(True)
        self.st_copyright_link.SetToolTip(
            wx.ToolTip(__copyright__)
        )
        self.st_copyright_link.UpdateLink()

        # Bottom
        self.sizer_btm_r.Add(
            self.sizer_btm_r_h, flag=wx.RIGHT | wx.ALIGN_RIGHT, border=4
        )
        self.sizer_btm_r.Add(
            self.st_copyright_link, flag=wx.RIGHT | wx.ALIGN_RIGHT, border=4
        )

        self.sizer_v.Add((-1, 4))
        self.sizer_v.Add(self.sizer_btm, proportion=0, flag=wx.EXPAND, border=4)
        self.sizer_btm.Add(
            self.sizer_btm_l, proportion=1, flag=wx.LEFT | wx.EXPAND, border=4
        )
        self.sizer_btm.Add(
            self.sizer_btm_r, proportion=1, flag=wx.RIGHT | wx.EXPAND, border=4
        )
        self.sizer_v.Add((-1, 4))

        if not x2l:
            self.cb_config.Hide()

        # Init panel
        self.panel.Layout()
        self.sizer_v.Hide(self.sizer_cfg_1, recursive=True)
        self.sizer_v.Hide(self.sizer_cfg_2, recursive=True)
        self.sizer_v.Hide(self.sizer_cfg_3, recursive=True)
        self.sizer_v.Hide(self.sizer_cfg_4, recursive=True)
        self.sizer_v.Hide(self.sizer_cfg_5, recursive=True)

    # 响应button事件
    def on_convert_click(self, event):
        """Convert button clicked."""
        if self.cb_config.IsChecked():
            self.hide_config()

        self.btn_convert.Disable()
        self.clear_log()
        x2l.run()
        self.btn_convert.Enable()

    def clear_log(self):
        """Clear the log."""
        self.logs.Clear()
        self.logs.Refresh()

    def save_config(self):
        """Save the config file."""
        # print('save_config')
        x2l.INPUT_FOLDER = self.tc1.GetValue()
        x2l.OUTPUT_GD_FOLDER = self.tc2.GetValue()
        x2l.OUTPUT_GD_NAME_TEMPLATE = self.tc3.GetValue()
        x2l.OUTPUT_CSV_FOLDER = self.tc4.GetValue()
        x2l.OUTPUT_CSV_NAME_TEMPLATE = self.tc5.GetValue()
        x2l.save_config()

    def load_config(self):
        """Load the config file."""
        # print('load_config')
        x2l.load_config()
        self.tc1.Clear()
        self.tc2.Clear()
        self.tc3.Clear()
        self.tc4.Clear()
        self.tc5.Clear()
        self.tc1.Refresh()
        self.tc2.Refresh()
        self.tc3.Refresh()
        self.tc4.Refresh()
        self.tc5.Refresh()
        self.tc1.write(x2l.INPUT_FOLDER)
        self.tc2.write(x2l.OUTPUT_GD_FOLDER)
        self.tc3.write(x2l.OUTPUT_GD_NAME_TEMPLATE)
        self.tc4.write(x2l.OUTPUT_CSV_FOLDER)
        self.tc5.write(x2l.OUTPUT_CSV_NAME_TEMPLATE)

    def on_config_checked(self, event):
        """Config checkbox checked."""
        self.toggle_config()

    def toggle_config(self):
        """Toggle the config."""
        if self.cb_config.IsChecked():
            self.show_config()
        else:
            self.hide_config()

    def show_config(self):
        """Show the config."""
        self.load_config()
        self.sizer_v.Show(self.sizer_cfg_1, recursive=True)
        self.sizer_v.Show(self.sizer_cfg_2, recursive=True)
        self.sizer_v.Show(self.sizer_cfg_3, recursive=True)
        self.sizer_v.Show(self.sizer_cfg_4, recursive=True)
        self.sizer_v.Show(self.sizer_cfg_5, recursive=True)
        self.panel.Layout()
        self.cb_config.SetValue(True)

    def hide_config(self):
        """Hide the config."""
        self.save_config()
        self.sizer_v.Hide(self.sizer_cfg_1, recursive=True)
        self.sizer_v.Hide(self.sizer_cfg_2, recursive=True)
        self.sizer_v.Hide(self.sizer_cfg_3, recursive=True)
        self.sizer_v.Hide(self.sizer_cfg_4, recursive=True)
        self.sizer_v.Hide(self.sizer_cfg_5, recursive=True)
        self.panel.Layout()
        self.cb_config.SetValue(False)

    def write(self, prefix, s):
        """Write log."""
        self.logs.WriteText("[")
        self.logs.BeginTextColour(self.prefix_color[prefix])
        self.logs.WriteText(prefix)
        self.logs.EndTextColour()
        self.logs.WriteText("]")
        self.logs.WriteText(f" {s}\n")

    def resource_path(self, relative_path):
        """Get absolute path to resource, works for dev and for PyInstaller"""
        try:
            # PyInstaller creates a temp folder and stores path in _MEIPASS
            base_path = sys._MEIPASS
        except Exception:
            base_path = os.path.abspath(".")
        return os.path.join(base_path, relative_path)


class App(wx.App):
    """App class."""
    def OnInit(self):
        """Init the app."""
        self.main_frame = MainFrame(None, "Excel to GDScript Convertor")
        self.main_frame.Show()
        if x2l:
            x2l.set_gui(self.main_frame)
        return True


def main():
    """Main function."""
    app = App()
    app.MainLoop()


if __name__ == "__main__":
    main()
