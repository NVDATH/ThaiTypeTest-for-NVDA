# thaiTypeTest/addon/globalPlugins/thaiTypeTest/__init__.py

import os
import sys
import wx
import re
import random
import addonHandler
import globalPluginHandler
import gui
import speech
import tones
from scriptHandler import script
from itertools import zip_longest
import string
import difflib

# Set up the library path first
lib_path = os.path.join(os.path.dirname(__file__), "lib")
if lib_path not in sys.path:
    sys.path.insert(0, lib_path)

# Import external and standard libraries
from pythainlp.tokenize import word_tokenize
import requests
from bs4 import BeautifulSoup

# Initialize add-on translations
addonHandler.initTranslation()


def clean_text(text):
    """Cleans text by removing blank lines, special characters, and BOM."""
    text = text.split("\n")
    text = [line.strip() for line in text]
    text = [line for line in text if line]
    text = "\n".join(text)
    text = text.replace(u'\ufeff', '') # Remove BOM
    # Use Python's 're' module to remove special characters
    return re.sub(r"[{}[\]()\*#<>]", "", text)

def fetch_lyrics(url):
    """Fetches and parses lyrics from Kapook, Siamzone, or Meemodel."""
    try:
        headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.36'}
        response = requests.get(url, headers=headers, timeout=10)
        response.raise_for_status()
        soup = BeautifulSoup(response.content, 'html.parser')

        raw_text = None
        if "kapook.com" in url:
            # Method 1: Find the styled separator divs (for modern layouts).
            separator_divs = soup.select("div[align='center'][style*='font-size:16px']")
            if len(separator_divs) >= 2:
                lyrics_parts = []
                start_node = separator_divs[0]
                end_node = separator_divs[1]
                for sibling in start_node.find_next_siblings():
                    if sibling == end_node:
                        break
                    text = sibling.get_text(separator='\n').strip()
                    if text:
                        lyrics_parts.append(text)
                if lyrics_parts:
                    raw_text = "\n".join(lyrics_parts)

            # Method 2 (Fallback for old table-based layouts): If Method 1 fails.
            if not raw_text:
                header_cell = soup.select_one("td.lyrics")
                if header_cell:
                    # The lyrics are in the <td> of the next <tr> sibling
                    header_row = header_cell.find_parent("tr")
                    if header_row:
                        lyrics_row = header_row.find_next_sibling("tr")
                        if lyrics_row:
                            lyrics_cell = lyrics_row.select_one("td")
                            if lyrics_cell:
                                raw_text = lyrics_cell.get_text(separator='\n').strip()

            # Method 3 (Final fallback): If both above methods fail.
            if not raw_text:
                lyrics_div = soup.select_one(".lyric p, .lyric")
                if lyrics_div:
                    raw_text = lyrics_div.get_text(separator='\n').strip()
        
        elif "siamzone.com" in url:
            # Logic for Siamzone
            lyrics_div = soup.select_one("div.has-text-centered-mobile.is-size-5-desktop")
            if lyrics_div:
                lyrics_parts = []
                found_karaoke_marker = False
                for element in lyrics_div.children:
                    text_content = ""
                    if isinstance(element, str):
                        text_content = element.strip()
                    else:
                        text_content = element.get_text(separator='\n').strip()
                    if "คาราโอเกะ" in text_content.lower() or "karaoke" in text_content.lower():
                        found_karaoke_marker = True
                        break
                    if not found_karaoke_marker and text_content:
                        lyrics_parts.append(text_content)
                raw_text = "\n".join(lyrics_parts).strip()

        elif "เพลง.meemodel.com" in url or "xn--72c9bva0i.meemodel.com" in url:
            # Logic for Meemodel
            lyrics_div = soup.select_one("div#lyric-lyric")
            if lyrics_div:
                raw_text = lyrics_div.get_text(separator='\n').strip()

        if raw_text:
            return clean_text(raw_text)
        return None

    except Exception as e:
        import logHandler
        logHandler.log.error(f"Failed to fetch lyrics from {url}", exc_info=True)
        return None
        
class TestDialog(wx.Dialog):
    """The main dialog for the Thai Type Test add-on."""
    def __init__(self, parent):
        dialog_title = "ทดสอบพิมพ์ภาษาไทย"
        super(TestDialog, self).__init__(parent, title=dialog_title)
        
        self.word_bank_general = []
        self.word_bank_hard = []
        self.MODES = {
            "พิมพ์คำ (ทั่วไป)": {"is_sentence": False, "source_files": ["sentence_th.txt", "lyrics_th.txt"]},
            "พิมพ์คำ (ยาก)": {"is_sentence": False, "source_files": ["literature_th.txt"]},
            "พิมพ์ประโยค": {"file": "sentence_th.txt", "is_sentence": True},
            "พิมพ์เนื้อเพลง": {"file": "lyrics_th.txt", "is_sentence": True},
            "พิมพ์วรรณกรรม": {"file": "literature_th.txt", "is_sentence": True},
        }
        self.load_all_data()

        if not self.word_bank_general and not self.word_bank_hard:
            wx.CallAfter(self.Destroy)
            return

        self.incorrect_pairs = []
        self.isRunning = False
        self.testDurationMinutes = 1
        self.elapsedTime = 0
        self.current_item_index = 0
        self.total_correct_words = 0
        self.total_incorrect_words = 0
        
        self.panel = wx.Panel(self)
        self.mainSizer = wx.BoxSizer(wx.VERTICAL)
        
        self.initialize_ui()
        self.update_ui_state()
        self.update_title()

        self.panel.SetSizer(self.mainSizer)
        self.Fit()
        self.Center()
        self.SetSize((550, 500))

        self.timer = wx.Timer(self)
        self.Bind(wx.EVT_TIMER, self.on_timer, self.timer)
        self.Bind(wx.EVT_CLOSE, self.on_close)

    def load_all_data(self):
        """Loads all datasets and creates word banks."""
        temp_word_bank_general = set()
        temp_word_bank_hard = set()

        for mode_name, mode_info in self.MODES.items():
            if "file" in mode_info:
                try:
                    file_path = os.path.join(os.path.dirname(__file__), "lib", mode_info["file"])
                    with open(file_path, "r", encoding="utf-8") as f:
                        data = [line.strip() for line in f if line.strip() and not line.strip().startswith("#")]
                    mode_info["dataset"] = data
                except FileNotFoundError:
                    mode_info["dataset"] = []
        
        for mode_name, mode_info in self.MODES.items():
            if mode_info.get("is_sentence") and mode_info.get("dataset"):
                tokenized_data = [word_tokenize(s, engine="newmm") for s in mode_info["dataset"]]
                flat_list = [word for sublist in tokenized_data for word in sublist]
                if mode_name == "พิมพ์วรรณกรรม":
                    temp_word_bank_hard.update(flat_list)
                else:
                    temp_word_bank_general.update(flat_list)

        self.word_bank_general = list(temp_word_bank_general)
        self.word_bank_hard = list(temp_word_bank_hard)

        if self.word_bank_general:
            self.MODES["พิมพ์คำ (ทั่วไป)"]["dataset"] = self.word_bank_general
        if self.word_bank_hard:
            self.MODES["พิมพ์คำ (ยาก)"]["dataset"] = self.word_bank_hard


    def initialize_ui(self):
        setupSizer = wx.BoxSizer(wx.HORIZONTAL)
        mode_label_text = "โหมด:"
        modeLabel = wx.StaticText(self.panel, label=mode_label_text)
        mode_choices = list(self.MODES.keys())
        self.modeChoice = wx.Choice(self.panel, choices=mode_choices)
        self.modeChoice.SetSelection(0)
        
        time_label_text = "ต้องการทดสอบกี่นาที:"
        timeLabel = wx.StaticText(self.panel, label=time_label_text)
        self.timeSpinCtrl = wx.SpinCtrl(self.panel, min=1, max=10, initial=1)
        
        setupSizer.Add(modeLabel, 0, wx.ALL | wx.ALIGN_CENTER_VERTICAL, 5)
        setupSizer.Add(self.modeChoice, 1, wx.ALL | wx.ALIGN_CENTER_VERTICAL, 5)
        setupSizer.AddSpacer(20)
        setupSizer.Add(timeLabel, 0, wx.ALL | wx.ALIGN_CENTER_VERTICAL, 5)
        setupSizer.Add(self.timeSpinCtrl, 1, wx.ALL | wx.ALIGN_CENTER_VERTICAL, 5)
        
        start_button_label = "เริ่ม (&S)"
        self.startButton = wx.Button(self.panel, label=start_button_label)
        self.startButton.SetDefault()

        self.typingTextCtrl = wx.TextCtrl(self.panel, style=wx.TE_PROCESS_ENTER, size=(-1, 40))
        
        self.resultsTextCtrl = wx.TextCtrl(self.panel, style=wx.TE_MULTILINE | wx.TE_READONLY | wx.TE_WORDWRAP)
        self.resultsTextCtrl.Hide()
        
        self.dynamicButtonSizer = wx.BoxSizer(wx.HORIZONTAL)
        self.addLyricsButton = wx.Button(self.panel, label="เพิ่มเนื้อเพลงจาก URL")
        self.editDataButton = wx.Button(self.panel, label="แก้ไขชุดข้อมูล")
        self.dynamicButtonSizer.Add(self.addLyricsButton, 1, wx.EXPAND | wx.ALL, 5)
        self.dynamicButtonSizer.Add(self.editDataButton, 1, wx.EXPAND | wx.ALL, 5)

        actionSizer = wx.BoxSizer(wx.HORIZONTAL)
        close_button_label = "ปิด (&C)"
        self.closeButton = wx.Button(self.panel, id=wx.ID_CANCEL, label=close_button_label)
        actionSizer.AddStretchSpacer()
        actionSizer.Add(self.closeButton, 0, wx.ALL, 5)
        actionSizer.AddStretchSpacer()
        
        self.mainSizer.Add(setupSizer, 0, wx.EXPAND | wx.ALL, 10)
        self.mainSizer.Add(self.dynamicButtonSizer, 0, wx.EXPAND | wx.LEFT | wx.RIGHT, 5)
        self.mainSizer.Add(self.startButton, 0, wx.EXPAND | wx.ALL, 10)
        self.mainSizer.Add(self.typingTextCtrl, 1, wx.EXPAND | wx.LEFT | wx.RIGHT | wx.BOTTOM, 10)
        self.mainSizer.Add(self.resultsTextCtrl, 1, wx.EXPAND | wx.LEFT | wx.RIGHT | wx.BOTTOM, 10)
        self.mainSizer.Add(actionSizer, 0, wx.EXPAND | wx.LEFT | wx.RIGHT, 5)
        
        # Bind events
        self.modeChoice.Bind(wx.EVT_CHOICE, self.on_mode_change)
        self.timeSpinCtrl.Bind(wx.EVT_KEY_DOWN, self.on_key_down_on_setup_controls)
        self.startButton.Bind(wx.EVT_BUTTON, self.on_start)
        self.typingTextCtrl.Bind(wx.EVT_TEXT_ENTER, self.on_enter_press)
        self.typingTextCtrl.Bind(wx.EVT_TEXT_PASTE, self.on_paste)
        self.addLyricsButton.Bind(wx.EVT_BUTTON, self.on_add_lyrics)
        
        # CRITICAL FIX: The missing line is added here.
        self.editDataButton.Bind(wx.EVT_BUTTON, self.on_edit_dataset)

    def on_key_down_on_setup_controls(self, event):
        if event.GetKeyCode() in (wx.WXK_RETURN, wx.WXK_NUMPAD_ENTER):
            if self.startButton.IsEnabled():
                self.on_start(None)
        else:
            event.Skip()

    def on_paste(self, event):
        tones.beep(200, 50)
        return

    def on_mode_change(self, event):
        self.update_title()
        selected_mode = self.modeChoice.GetStringSelection()
        is_lyrics_mode = (selected_mode == "พิมพ์เนื้อเพลง")
        is_editable_mode = selected_mode in ["พิมพ์ประโยค", "พิมพ์เนื้อเพลง", "พิมพ์วรรณกรรม"]
        self.addLyricsButton.Show(is_lyrics_mode)
        self.editDataButton.Show(is_editable_mode)
        self.dynamicButtonSizer.Show(is_lyrics_mode or is_editable_mode)
        self.panel.Layout()

    def on_add_lyrics(self, event):
        clipboard = wx.TheClipboard
        if clipboard.Open():
            data = wx.TextDataObject()
            success = clipboard.GetData(data)
            clipboard.Close()
            if success:
                url = data.GetText()
                if "kapook.com" in url or "siamzone.com" in url or "เพลง.meemodel.com" in url or "xn--72c9bva0i.meemodel.com" in url:
                    speech.speakMessage("กำลังดึงข้อมูลเนื้อเพลง กรุณารอสักครู่")
                    lyrics = fetch_lyrics(url)
                    if lyrics:
                        file_path = os.path.join(os.path.dirname(__file__), "lib", "lyrics_th.txt")
                        try:
                            with open(file_path, "a", encoding="utf-8") as f:
                                f.write(f"\n#credit: {url}\n")
                                f.write(lyrics)
                            gui.messageBox("เพิ่มเนื้อเพลงเรียบร้อยแล้ว", "สำเร็จ", wx.OK | wx.ICON_INFORMATION)
                            self.load_all_data()
                        except Exception as e:
                             gui.messageBox(f"ไม่สามารถบันทึกไฟล์เนื้อเพลงได้: {e}", "ข้อผิดพลาด", wx.OK | wx.ICON_ERROR)
                    else:
                        gui.messageBox("ไม่สามารถดึงเนื้อเพลงจาก URL ที่ให้มาได้", "ล้มเหลว", wx.OK | wx.ICON_ERROR)
                else:
                    self.ask_to_open_file("URL ไม่ถูกต้องหรือไม่รองรับ", "lyrics_th.txt")
            else:
                self.ask_to_open_file("ไม่พบ URL ใน Clipboard", "lyrics_th.txt")
    
    def ask_to_open_file(self, message, filename):
        dialog = wx.MessageDialog(self, f"{message}\n\nคุณต้องการเปิดไฟล์ {filename} เพื่อแก้ไขด้วยตนเองหรือไม่?", "แจ้งเตือน", wx.YES_NO | wx.ICON_QUESTION)
        if dialog.ShowModal() == wx.ID_YES:
            self.open_data_file(filename)
        dialog.Destroy()
    
    def on_edit_dataset(self, event):
        selected_mode = self.modeChoice.GetStringSelection()
        filename = self.MODES[selected_mode].get("file")
        if filename:
            self.open_data_file(filename)

    def open_data_file(self, filename):
        try:
            file_path = os.path.join(os.path.dirname(__file__), "lib", filename)
            if not os.path.exists(file_path):
                open(file_path, 'a').close()
            os.startfile(file_path)
        except Exception as e:
            gui.messageBox(f"ไม่สามารถเปิดไฟล์ได้: {e}", "ข้อผิดพลาด", wx.OK | wx.ICON_ERROR)

    def update_title(self, event=None):
        base_title = "ทดสอบพิมพ์ภาษาไทย"
        if self.isRunning and hasattr(self, 'current_dataset') and self.current_item_index < len(self.current_dataset):
            current_item = self.current_dataset[self.current_item_index]
            self.SetTitle(current_item)
        else:
            mode_text = self.modeChoice.GetStringSelection()
            new_title = f"{base_title} - [{mode_text}]"
            self.SetTitle(new_title)

    def update_ui_state(self):
        is_setting_up = not self.isRunning
        self.modeChoice.Enable(is_setting_up)
        self.timeSpinCtrl.Enable(is_setting_up)
        self.startButton.Enable(is_setting_up)
        self.typingTextCtrl.Show(self.isRunning)
        self.typingTextCtrl.Enable(self.isRunning)
        if not self.isRunning:
            self.update_title()
            self.on_mode_change(None)
        if self.isRunning:
            self.resultsTextCtrl.Hide()
            self.addLyricsButton.Hide()
            self.editDataButton.Hide()
            self.dynamicButtonSizer.Show(False)
        self.panel.Layout()
    
    def on_start(self, event):
        self.load_all_data()
        selected_mode = self.modeChoice.GetStringSelection()
        if not self.MODES[selected_mode].get("dataset"):
            gui.messageBox(f"ไม่พบชุดข้อมูลสำหรับโหมด '{selected_mode}'\nกรุณาเพิ่มข้อมูลในไฟล์ .txt หรือเลือกโหมดอื่น", "ข้อผิดพลาด", wx.OK | wx.ICON_ERROR)
            return
        self.isRunning = True
        self.update_ui_state()
        self.typingTextCtrl.SetFocus()
        selected_time = self.timeSpinCtrl.GetValue()
        warning_message = f"กำลังจะทดสอบโหมด '{selected_mode}' ในเวลา {selected_time} นาที กรุณาตรวจสอบว่าได้เปลี่ยนแป้นพิมพ์เป็นภาษาไทยไว้แล้ว"
        speech.speakMessage(warning_message)
        wx.CallLater(5000, self.begin_test_logic)
    
    def begin_test_logic(self):
        if not self.IsShown(): return
        self.elapsedTime = 0
        self.current_item_index = 0
        self.total_correct_words = 0
        self.total_incorrect_words = 0
        self.incorrect_pairs = []
        self.testDurationMinutes = self.timeSpinCtrl.GetValue()
        selected_mode = self.modeChoice.GetStringSelection()
        self.current_dataset = list(self.MODES[selected_mode].get("dataset", []))
        if not self.current_dataset:
            self.isRunning = False
            self.update_ui_state()
            return
        random.shuffle(self.current_dataset)
        self.typingTextCtrl.Clear()
        tones.beep(1000, 100)
        self.timer.Start(1000)
        self.speak_current_item()

    def on_enter_press(self, event):
        if not self.isRunning: return
        typed_item = self.typingTextCtrl.GetValue().strip()
        if not typed_item:
            self.speak_current_item()
            return
        correct_item = self.current_dataset[self.current_item_index]
        is_sentence_mode = self.MODES[self.modeChoice.GetStringSelection()].get("is_sentence", False)
        if not is_sentence_mode:
            if typed_item == correct_item:
                self.total_correct_words += 1
            else:
                self.total_incorrect_words += 1
                self.incorrect_pairs.append((correct_item, typed_item))
        else:
            punctuation_to_ignore = string.punctuation + "ๆฯ“”"
            correct_tokens = word_tokenize(correct_item, engine="newmm")
            typed_tokens = word_tokenize(typed_item, engine="newmm")
            correct_words_filtered = [word for word in correct_tokens if word not in punctuation_to_ignore and not word.isspace()]
            typed_words_filtered = [word for word in typed_tokens if word not in punctuation_to_ignore and not word.isspace()]
            matcher = difflib.SequenceMatcher(None, correct_words_filtered, typed_words_filtered)
            sentence_correct, sentence_incorrect = 0, 0
            for tag, i1, i2, j1, j2 in matcher.get_opcodes():
                if tag == 'equal':
                    sentence_correct += (i2 - i1)
                else:
                    num_correct_in_chunk = i2 - i1
                    num_typed_in_chunk = j2 - j1
                    sentence_incorrect += max(num_correct_in_chunk, num_typed_in_chunk)
            self.total_correct_words += sentence_correct
            self.total_incorrect_words += sentence_incorrect
            if sentence_incorrect > 0:
                self.incorrect_pairs.append((correct_item, typed_item))
        self.typingTextCtrl.Clear()
        self.current_item_index += 1
        if self.current_item_index < len(self.current_dataset):
            self.speak_current_item()
        else:
            self.end_test()

    def on_timer(self, event):
        try:
            if not self.IsShown():
                event.GetTimer().Stop()
                return
        except wx.wxAssertionError:
            event.GetTimer().Stop()
            return
        self.elapsedTime += 1
        total_duration = self.testDurationMinutes * 60
        time_remaining = total_duration - self.elapsedTime
        if time_remaining == 4: tones.beep(440, 70)
        elif time_remaining == 3: tones.beep(550, 70)
        elif time_remaining == 2: tones.beep(660, 70)
        elif time_remaining == 1: tones.beep(770, 70)
        if self.elapsedTime > 0 and self.elapsedTime % 60 == 0:
            tones.beep(440, 50)
        if self.elapsedTime >= total_duration:
            self.end_test()

    def on_close(self, event):
        self.timer.Stop()
        self.Destroy()

    def speak_current_item(self):
        if self.isRunning and self.current_item_index < len(self.current_dataset):
            self.update_title()
            speech.speakMessage(self.current_dataset[self.current_item_index])

    def end_test(self):
        self.timer.Stop()
        self.isRunning = False
        tones.beep(880, 500)
        gui.messageBox("การทดสอบสิ้นสุดแล้ว", "สิ้นสุดการทดสอบ", wx.OK | wx.ICON_INFORMATION)
        total_words_typed = self.total_correct_words + self.total_incorrect_words
        accuracy = (self.total_correct_words / total_words_typed) * 100 if total_words_typed > 0 else 0
        net_wpm = self.total_correct_words / self.testDurationMinutes if self.testDurationMinutes > 0 else 0
        gross_wpm = total_words_typed / self.testDurationMinutes if self.testDurationMinutes > 0 else 0
        
        summary = (
            f"สรุปผล:\n"
            f"- ความเร็วรวม (Gross WPM): {gross_wpm:.1f} คำต่อนาที\n"
            f"- ความเร็วสุทธิ (Net WPM): {net_wpm:.1f} คำต่อนาที\n"
            f"- ความแม่นยำ: {accuracy:.1f}%\n"
            f"- พิมพ์ถูกทั้งหมด: {self.total_correct_words} คำ\n"
            f"- พิมพ์ผิดทั้งหมด: {self.total_incorrect_words} คำ\n"
        )
        
        details = ""
        if self.incorrect_pairs:
            selected_mode = self.modeChoice.GetStringSelection()
            unit = "คำ" if not self.MODES[selected_mode].get("is_sentence") else "ประโยค"
            details += f"\n----------\n{unit}ที่พิมพ์ผิด:\n"
            for correct, typed in self.incorrect_pairs:
                details += f"- ต้นฉบับ: {correct}\n"
                details += f"- ที่คุณพิมพ์: {typed}\n\n"
        
        standard_note = "\nเกณฑ์มาตรฐาน: โดยทั่วไปคะแนน Net WPM ที่น่าเชื่อถือควรมีความแม่นยำตั้งแต่ 95% ขึ้นไป"
        full_report = summary + details.strip() + standard_note
        self.resultsTextCtrl.SetValue(full_report)
        self.update_ui_state()
        self.resultsTextCtrl.Show()
        self.panel.Layout()
        self.Fit()
        self.resultsTextCtrl.SetFocus()
        
class GlobalPlugin(globalPluginHandler.GlobalPlugin):
    def __init__(self):
        super(GlobalPlugin, self).__init__()
        self.menu_item = None
        wx.CallLater(1, self.add_menu_item)

    def add_menu_item(self):
        try:
            tools_menu = gui.mainFrame.sysTrayIcon.toolsMenu
            menu_text = "ทดสอบพิมพ์ภาษาไทย..."
            help_text = "เปิดหน้าต่างเพื่อทดสอบความเร็วในการพิมพ์ภาษาไทย"
            self.menu_item = tools_menu.Append(wx.ID_ANY, menu_text, help_text)
            gui.mainFrame.sysTrayIcon.Bind(wx.EVT_MENU, self.on_show_dialog_menu, self.menu_item)
        except Exception as e:
            import logHandler
            logHandler.log.error("Failed to add Thai Type Test menu item", exc_info=True)

    def show_dialog(self):
        for child in gui.mainFrame.GetChildren():
            if isinstance(child, TestDialog):
                child.Raise()
                return
        dialog = TestDialog(gui.mainFrame.prePopup())
        dialog.Show()

    def on_show_dialog_menu(self, event):
        self.show_dialog()

    @script(
        description="เปิดหน้าต่างทดสอบพิมพ์ภาษาไทย",
        category="Thai Type Test"
    )
    def script_showDialog(self, gesture):
        wx.CallLater(1, self.show_dialog)

    def terminate(self):
        try:
            tools_menu = gui.mainFrame.sysTrayIcon.toolsMenu
            gui.mainFrame.sysTrayIcon.Unbind(wx.EVT_MENU, handler=self.on_show_dialog_menu, source=self.menu_item)
            tools_menu.Remove(self.menu_item)
        except Exception:
            pass
            
class GlobalPlugin(globalPluginHandler.GlobalPlugin):
    """The Global Plugin to integrate the add-on into NVDA."""
    def __init__(self):
        super(GlobalPlugin, self).__init__()
        self.menu_item = None
        wx.CallLater(1, self.add_menu_item)

    def add_menu_item(self):
        try:
            tools_menu = gui.mainFrame.sysTrayIcon.toolsMenu
            menu_text = "ทดสอบพิมพ์ภาษาไทย..."
            help_text = "เปิดหน้าต่างเพื่อทดสอบความเร็วในการพิมพ์ภาษาไทย"
            self.menu_item = tools_menu.Append(wx.ID_ANY, menu_text, help_text)
            gui.mainFrame.sysTrayIcon.Bind(wx.EVT_MENU, self.on_show_dialog_menu, self.menu_item)
        except Exception as e:
            import logHandler
            logHandler.log.error("Failed to add Thai Type Test menu item", exc_info=True)

    def show_dialog(self):
        for child in gui.mainFrame.GetChildren():
            if isinstance(child, TestDialog):
                child.Raise()
                return
        dialog = TestDialog(gui.mainFrame.prePopup())
        dialog.Show()

    def on_show_dialog_menu(self, event):
        self.show_dialog()

    @script(
        description="เปิดหน้าต่างทดสอบพิมพ์ภาษาไทย",
        category="Thai Type Test"
    )
    def script_showDialog(self, gesture):
        wx.CallLater(1, self.show_dialog)

    def terminate(self):
        try:
            tools_menu = gui.mainFrame.sysTrayIcon.toolsMenu
            gui.mainFrame.sysTrayIcon.Unbind(wx.EVT_MENU, handler=self.on_show_dialog_menu, source=self.menu_item)
            tools_menu.Remove(self.menu_item)
        except Exception:
            pass