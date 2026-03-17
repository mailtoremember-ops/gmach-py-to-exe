import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
import csv
from datetime import datetime, date
import ttkbootstrap as ttkb
import ttkbootstrap.toast as toast
import os
import copy
import tkinter.font as tkFont

# הגדרות קבצי נתונים המשמשים לטעינה ושמירה של נתוני היישום
TOOLS_FILE = 'tools.csv'
BORROWERS_FILE = 'borrowers.csv'
BORROWING_HISTORY_FILE = 'borrowing_history.csv'
BORROWING_HISTORY_FIELDNAMES = ["תאריך השאלה", "שם הכלי", "שם השואל"] # שדות נתוני היסטוריית השאלות

class DataManager:
    """
    מחלקה לניהול טעינה ושמירה של נתוני כלים, שואלים והיסטוריית השאלות.
    """
    def __init__(self):
        self.tools_data = []
        self.borrowers_data = []
        self.borrowing_history_data = []
        self.history = [] # ישמור את המצב ההתחלתי: (tools_copy, borrowers_copy, initial_history_len)

        self.load_all_data()
        self.save_state() # שמור מצב התחלתי לאחר טעינה

    def load_data(self, filename):
        """טוען נתונים מקובץ CSV."""
        data = []
        try:
            with open(filename, mode='r', newline='', encoding='utf-8') as file:
                reader = csv.DictReader(file)
                data = list(reader)
        except FileNotFoundError:
            pass # קובץ לא נמצא, יתחיל עם נתונים ריקים.
        except Exception as e:
            print(f"Error loading {filename}: {e}") # הדפסה לצרכי דיבוג
            pass
        return data

    def save_data(self, filename, data, fieldnames):
        """שומר נתונים לקובץ CSV."""
        try:
            with open(filename, mode='w', newline='', encoding='utf-8') as file:
                writer = csv.DictWriter(file, fieldnames=fieldnames)
                writer.writeheader()
                writer.writerows(data)
        except Exception as e:
             print(f"Error saving {filename}: {e}") # הדפסה לצרכי דיבוג
             pass

    def load_all_data(self):
        """טוען את כל הנתונים בתחילת היישום."""
        self.tools_data = self.load_data(TOOLS_FILE)
        # הוספת שדות חסרים עם ערכי ברירת מחדל אם לא קיימים בקובץ (הסרת 'מיקום')
        default_tool = {
            "שם הכלי": "", "תיאור כלי": "", "סטטוס": "זמין", "שם השואל": "",
            "תאריך השאלה": "", "מונה השאלות": "0", "מספר סידורי": "" # 'מיקום' הוסר
            }
        for tool in self.tools_data:
            # ניקוי שדה 'מיקום' מטעינת קבצים ישנים
            if 'מיקום' in tool:
                del tool['מיקום']
            for key, default_value in default_tool.items():
                if key not in tool:
                    tool[key] = default_value
                if key == "מונה השאלות": # לוודא שמונה השאלות הוא מחרוזת של מספר
                     try:
                         tool[key] = str(int(tool[key]))
                     except (ValueError, TypeError):
                          tool[key] = "0"

        self.borrowers_data = self.load_data(BORROWERS_FILE)
        self.borrowing_history_data = self.load_data(BORROWING_HISTORY_FILE)

    def save_all_data(self):
        """שומר את כל הנתונים חזרה לקבצי CSV."""
        # הסרת 'מיקום' מרשימת השדות לכלי
        tool_fieldnames = ["שם הכלי", "תיאור כלי", "סטטוס", "שם השואל", "תאריך השאלה", "מונה השאלות", "מספר סידורי"]
        borrower_fieldnames = ["שם השואל", "מספר טלפון", "כתובת"]
        self.save_data(TOOLS_FILE, self.tools_data, tool_fieldnames)
        self.save_data(BORROWERS_FILE, self.borrowers_data, borrower_fieldnames)
        self.save_data(BORROWING_HISTORY_FILE, self.borrowing_history_data, BORROWING_HISTORY_FIELDNAMES)

    # --- פונקציות לניהול היסטוריה (Undo/Reset) ---

    def save_state(self):
        """שומר את המצב ההתחלתי של הנתונים להיסטוריה בפעם הראשונה."""
        if not self.history:
            initial_history_length = len(self.borrowing_history_data)
            current_state_tuple = (
                copy.deepcopy(self.tools_data),
                copy.deepcopy(self.borrowers_data),
                initial_history_length
            )
            self.history.append(current_state_tuple)

    def restore_initial_state(self):
        """מחזיר את מצב הנתונים למצב ההתחלתי בעת פתיחת התוכנה."""
        if len(self.history) > 0:
            initial_state_tuple = self.history[0]

            self.tools_data = copy.deepcopy(initial_state_tuple[0])
            self.borrowers_data = copy.deepcopy(initial_state_tuple[1])

            initial_history_length = initial_state_tuple[2]
            self.borrowing_history_data = self.borrowing_history_data[:initial_history_length]

            self.history = [copy.deepcopy(initial_state_tuple)]
            return True
        return False

    def append_borrowing_history(self, tool_name, borrower_name, borrow_date):
         """מוסיף רשומת השאלה להיסטוריה בזיכרון."""
         record = {
             BORROWING_HISTORY_FIELDNAMES[0]: borrow_date,
             BORROWING_HISTORY_FIELDNAMES[1]: tool_name,
             BORROWING_HISTORY_FIELDNAMES[2]: borrower_name
         }
         self.borrowing_history_data.append(record)

    # --- פונקציות עזר לחישוב ימים ---
    def calculate_days_borrowed(self, borrow_date_str):
        """מחשב את מספר הימים שעברו מתאריך ההשאלה."""
        if not borrow_date_str:
            return ""
        try:
            borrow_date = datetime.strptime(borrow_date_str, '%Y-%m-%d').date()
            today = date.today()
            delta = today - borrow_date
            return str(delta.days)
        except (ValueError, TypeError):
            return "שגיאת תאריך"

    # --- פונקציות למציאת אובייקטים בנתונים ---
    def find_tool(self, name, description):
        """מוצא כלי ברשימת tools_data לפי שם ותיאור."""
        for tool in self.tools_data:
            tool_name_in_data = tool.get('שם הכלי')
            tool_description_in_data = tool.get('תיאור כלי', '')
            if tool_name_in_data == name and tool_description_in_data == description:
                return tool
        return None

    def find_tool_by_serial(self, serial_number):
        """מוצא כלי ברשימת tools_data לפי מספר סידורי."""
        if not serial_number:
            return None
        for tool in self.tools_data:
             if tool.get("מספר סידורי", "") == serial_number:
                  return tool
        return None


    def find_borrower(self, name):
        """מוצא שואל ברשימת borrowers_data לפי שם."""
        for borrower in self.borrowers_data:
            if borrower.get('שם השואל') == name:
                return borrower
        return None


class ToolBorrowingApp:
    """
    המחלקה הראשית של היישום לניהול השאלת כלי עבודה.
    אחראית על בניית הממשק הגרפי, ניהול אירועים וקישור ל-DataManager.
    """
    def __init__(self, root):
        self.root = root
        self.data_manager = DataManager()

        # אתחול משתני ממשק המשתמש הראשי
        self.search_entry = None
        self.tools_tree = None
        self.borrower_name_entry = None
        self.borrower_phone_entry = None
        self.borrower_address_entry = None
        self.borrower_tree = None
        self.borrowed_tools_tree = None

        # --- שורות חדשות שיש להוסיף לאתחול ---
        # אתחול משתני חלונות הניהול ל-None
        self.tool_management_window = None
        self.borrower_management_window = None
        # -------------------------------------
        self.setup_ui() # זו המתודה שיוצרת ומקצה את הווידג'טים בפועל למשתנים לעיל

        self.filter_tools(None)
        self.filter_borrower_table(None)

    def show_toast(self, message: str, title: str = "הודעת מערכת", duration: int = 3000, bootstyle: str = "info", position: tuple = (0, 250, 'n'), icon: str = ""):
        """
        מציגה הודעת "טוסט" זמנית על המסך במיקום שצויין.
        """
        t = toast.ToastNotification(
            title=title,
            message=message,
            duration=duration,
            bootstyle=bootstyle,
            position=position,
            icon = ""
        )
        t.show_toast()

    # --- פונקציה לניקוי שדה קלט באמצעות Escape ---
    def clear_entry_on_escape(self, event):
        """מנקה את תוכן השדה שקיבל את האירוע Escape ומפעיל סינון מחדש אם רלוונטי."""
        if event.keysym != 'Escape':
            return

        entry_widget = event.widget

        if entry_widget is self.search_entry:
            entry_widget.delete(0, tk.END)
            self.filter_tools(None)
        elif entry_widget in [self.borrower_name_entry, self.borrower_phone_entry, self.borrower_address_entry]:
             self.borrower_name_entry.delete(0, tk.END)
             self.borrower_phone_entry.delete(0, tk.END)
             self.borrower_address_entry.delete(0, tk.END)
             self.filter_borrower_table(None)
             self.borrower_tree.selection_set([])
             self.refresh_borrowed_tools_table([])

        # הוספת ניקוי עבור שדות בחלון ניהול הכלים אם הם קיימים (רק אם חלון הניהול פתוח)
        # בדיקה מתוקנת: ודא שהמשתנה אינו None לפני קריאה ל winfo_exists()
        if self.tool_management_window and self.tool_management_window.winfo_exists():
             if entry_widget in [self.tool_management_window.tool_name_entry,
                                  self.tool_management_window.tool_description_entry,
                                  self.tool_management_window.serial_number_entry]: # 'מיקום' הוסר
                   # נקה את כל שדות ניהול הכלים יחד
                   self.tool_management_window.tool_name_entry.delete(0, tk.END)
                   self.tool_management_window.tool_description_entry.delete(0, tk.END)
                   self.tool_management_window.serial_number_entry.delete(0, tk.END) # 'מיקום' הוסר
                   # בטל בחירה בטבלת ניהול הכלים
                   self.tool_management_window.tool_list_tree.selection_set([])
                   # אפשר מחדש את כפתור הוספה, נטרל עריכה/מחיקה
                   self.on_tool_list_select_window(None, self.tool_management_window) # שימוש בפונקציה הקיימת לאיפוס מצב הכפתורים

        # הוספת ניקוי עבור שדות בחלון ניהול השואלים אם הם קיימים (תיקון זהה)
        if self.borrower_management_window and self.borrower_management_window.winfo_exists():
            if entry_widget in [self.borrower_management_window.borrower_name_entry,
                                self.borrower_management_window.borrower_phone_entry,
                                self.borrower_management_window.borrower_address_entry]:
                self.borrower_management_window.borrower_name_entry.delete(0, tk.END)
                self.borrower_management_window.borrower_phone_entry.delete(0, tk.END)
                self.borrower_management_window.borrower_address_entry.delete(0, tk.END)


        return 'break'


    def setup_ui(self):
        """בניית הממשק הגרפי של החלון הראשי."""
        self.root.option_add('*TCombobox*Listbox.font', ('Helvetica', 12))
        self.root.option_add('*TCombobox*Listbox.background', 'white')
        self.root.option_add('*font', ('Helvetica', 12))
        self.root.title("השאלת כלי עבודה - גמ\"ח")
        self.root.geometry("1185x800")

        self.root.protocol("WM_DELETE_WINDOW", lambda: (self.data_manager.save_all_data(), self.root.destroy()))

        self.top_control_frame = ttkb.Frame(self.root, padding="15")
        self.top_control_frame.grid(row=0, column=0, sticky="ew", padx=10, pady=(5, 0))

        search_label = ttkb.Label(self.top_control_frame, text="חיפוש", font=('Helvetica', 12))
        search_label.grid(row=0, column=6, sticky="e", padx=(2, 2), pady=(2, 2))

        self.search_entry = ttkb.Entry(self.top_control_frame, width=40, justify='right')
        self.search_entry.grid(row=0, column=5, sticky="ew", padx=5, pady=(10, 10))
        self.search_entry.bind("<KeyRelease>", self.filter_tools)
        self.search_entry.bind("<Escape>", self.clear_entry_on_escape)


        statistics_button = ttkb.Button(self.top_control_frame, text="סטטיסטיקה", bootstyle="info", command=self.show_statistics)
        statistics_button.grid(row=0, column=4, sticky="w", padx=5, pady=(10, 10))

        history_button = ttkb.Button(self.top_control_frame, text="היסטוריית השאלות", bootstyle="secondary", command=self.show_borrowing_history)
        history_button.grid(row=0, column=3, sticky="w", padx=5, pady=(10, 10))

        undo_button = ttkb.Button(self.top_control_frame, text="ביטול פעולות", bootstyle="danger", command=self.undo_last_action)
        undo_button.grid(row=0, column=2, sticky="w", padx=5, pady=(10, 10))

        manage_tools_button = ttkb.Button(self.top_control_frame, text="ניהול כלים", bootstyle="primary", command=self.show_add_tool_window)
        manage_tools_button.grid(row=0, column=0, sticky="w", padx=5, pady=(10, 10))

        manage_borrowers_button = ttkb.Button(self.top_control_frame, text="ניהול שואלים", bootstyle="primary", command=self.show_manage_borrowers_window)
        manage_borrowers_button.grid(row=0, column=1, sticky="w", padx=5, pady=(10, 10))

        self.top_control_frame.grid_columnconfigure(4, weight=1)

        self.tools_frame = ttkb.Frame(self.root, padding="15")
        self.tools_frame.grid(row=1, column=0, sticky="nsew", padx=10, pady=(0, 5))

        self.root.grid_rowconfigure(1, weight=1)
        self.root.grid_columnconfigure(0, weight=1)

        # הגדרת העמודות שיוצגו בטבלה הראשית (לפי סדר לוגי) - 'מיקום' לא היה כאן, נשאר ללא שינוי
        tools_tree_columns_logical = ("ימים", "שם השואל", "סטטוס", "תיאור כלי", "שם הכלי", "מספר סידורי")

        self.tools_tree = ttk.Treeview(self.tools_frame, columns=tools_tree_columns_logical, show="headings")
        tools_tree_style = ttk.Style()
        tools_tree_style.configure("Treeview.Heading", font=('Helvetica', 12, 'bold'))
        tools_tree_style.configure("Treeview", font=('Helvetica', 12), rowheight=25)

        self.tools_tree.tag_configure('borrowed', background='#ffe0e0')

        # הגדרת כותרות ורוחב לעמודות המוצגות (סדר ויזואלי הפוך מהלוגי) - נשאר ללא שינוי
        self.tools_tree.heading("מספר סידורי", text="#", anchor="e")
        self.tools_tree.column("מספר סידורי", width=10, anchor="e")
        self.tools_tree.heading("שם הכלי", text="שם הכלי", anchor="e")
        self.tools_tree.column("שם הכלי", width=150, anchor="e")
        self.tools_tree.heading("תיאור כלי", text="תיאור כלי", anchor="e")
        self.tools_tree.column("תיאור כלי", width=200, anchor="e")
        self.tools_tree.heading("סטטוס", text="סטטוס", anchor="e")
        self.tools_tree.column("סטטוס", width=100, anchor="e")
        self.tools_tree.heading("שם השואל", text="שם השואל", anchor="e")
        self.tools_tree.column("שם השואל", width=150, anchor="e")
        self.tools_tree.heading("ימים", text="ימים", anchor="e")
        self.tools_tree.column("ימים", width=70, anchor="e")


        tools_tree_scrollbar = ttkb.Scrollbar(self.tools_frame, orient="vertical", command=self.tools_tree.yview)
        self.tools_tree.configure(yscrollcommand=tools_tree_scrollbar.set)

        tools_tree_scrollbar.grid(row=0, column=1, sticky="ns")
        self.tools_tree.grid(row=0, column=0, sticky="nsew")

        self.tools_frame.grid_columnconfigure(0, weight=1)
        self.tools_frame.grid_rowconfigure(0, weight=1)

        self.tools_tree.bind("<Button-3>", self.show_tools_context_menu)
        # הסרת הקישור לחיצה כפולה שהציג את המיקום
        # self.tools_tree.bind("<Double-1>", self.show_tool_details)


        self.add_borrower_and_search_frame = ttkb.Frame(self.root, padding="15")
        self.add_borrower_and_search_frame.grid(row=2, column=0, sticky="ew", padx=10, pady=(5, 0))

        borrower_name_label = ttkb.Label(self.add_borrower_and_search_frame, text="שם")
        borrower_name_label.grid(row=0, column=6, sticky="e", padx=(2, 2))

        self.borrower_name_entry = ttkb.Entry(self.add_borrower_and_search_frame, width=19, justify='right')
        self.borrower_name_entry.grid(row=0, column=5, sticky="ew", padx=5)
        self.borrower_name_entry.bind("<Escape>", self.clear_entry_on_escape)


        borrower_phone_label = ttkb.Label(self.add_borrower_and_search_frame, text="טלפון")
        borrower_phone_label.grid(row=0, column=4, sticky="e", padx=(2, 2))

        self.borrower_phone_entry = ttkb.Entry(self.add_borrower_and_search_frame, width=8, justify='right')
        self.borrower_phone_entry.grid(row=0, column=3, sticky="ew", padx=5)
        self.borrower_phone_entry.bind("<Escape>", self.clear_entry_on_escape)


        borrower_address_label = ttkb.Label(self.add_borrower_and_search_frame, text="כתובת")
        borrower_address_label.grid(row=0, column=2, sticky="e", padx=(2, 2))

        self.borrower_address_entry = ttkb.Entry(self.add_borrower_and_search_frame, width=10, justify='right')
        self.borrower_address_entry.grid(row=0, column=1, sticky="ew", padx=5)
        self.borrower_address_entry.bind("<Escape>", self.clear_entry_on_escape)


        borrower_buttons_frame = ttkb.Frame(self.add_borrower_and_search_frame)
        borrower_buttons_frame.grid(row=0, column=0, sticky="w", padx=5)

        add_borrower_button = ttkb.Button(borrower_buttons_frame, text="הוסף שואל", bootstyle="success", command=self.add_borrower)
        add_borrower_button.pack(side="right", padx=5)

        borrow_tool_button = ttkb.Button(borrower_buttons_frame, text="השאל כלי ", bootstyle="primary", command=self.borrow_tool)
        borrow_tool_button.pack(side="right", padx=5)

        return_tool_button = ttkb.Button(borrower_buttons_frame, text="החזר כלי ", bootstyle="warning", command=self.return_tool)
        return_tool_button.pack(side="right", padx=5)

        self.add_borrower_and_search_frame.grid_columnconfigure(1, weight=1)
        self.add_borrower_and_search_frame.grid_columnconfigure(3, weight=1)
        self.add_borrower_and_search_frame.grid_columnconfigure(5, weight=1)

        self.borrower_name_entry.bind("<KeyRelease>", self.filter_borrower_table)
        self.borrower_phone_entry.bind("<KeyRelease>", self.filter_borrower_table)
        self.borrower_address_entry.bind("<KeyRelease>", self.filter_borrower_table)

        self.borrower_table_frame = ttkb.Frame(self.root, padding="15")
        self.borrower_table_frame.grid(row=3, column=0, sticky="ew", padx=10, pady=(0, 2))

        borrower_tree_columns_logical = ("כתובת", "מספר טלפון", "שם השואל")

        self.borrower_tree = ttk.Treeview(self.borrower_table_frame, columns=borrower_tree_columns_logical, show="headings", height=4)
        borrower_tree_style = ttk.Style()
        borrower_tree_style.configure("Treeview.Heading", font=('Helvetica', 12, 'bold'))
        borrower_tree_style.configure("Treeview", font=('Helvetica', 12), rowheight=25)

        self.borrower_tree.heading("שם השואל", text="שם השואל", anchor="e")
        self.borrower_tree.column("שם השואל", width=150, anchor="e")
        self.borrower_tree.heading("מספר טלפון", text="מספר טלפון", anchor="e")
        self.borrower_tree.column("מספר טלפון", width=150, anchor="e")
        self.borrower_tree.heading("כתובת", text="כתובת", anchor="e")
        self.borrower_tree.column("כתובת", width=150, anchor="e")

        borrower_tree_scrollbar = ttkb.Scrollbar(self.borrower_table_frame, orient="vertical", command=self.borrower_tree.yview)
        self.borrower_tree.configure(yscrollcommand=borrower_tree_scrollbar.set)

        borrower_tree_scrollbar.grid(row=0, column=1, sticky="ns")
        self.borrower_tree.grid(row=0, column=0, sticky="nsew")

        self.borrower_table_frame.grid_columnconfigure(0, weight=1)

        self.borrower_tree.bind("<<TreeviewSelect>>", self.on_borrower_table_select)

        self.borrowed_tools_frame = ttkb.Frame(self.root, padding=(15, 0, 15, 15))
        self.borrowed_tools_frame.grid(row=4, column=0, sticky="nsew", padx=10, pady=(0, 10))

        self.root.grid_rowconfigure(4, weight=1)

        borrowed_tools_label = ttkb.Label(self.borrowed_tools_frame, text="כלים מושאלים", font=('Helvetica', 12, 'bold'))
        borrowed_tools_label.grid(row=0, column=0, sticky="e", pady=(0, 5))

        borrowed_tools_tree_columns_logical = ("תאריך השאלה", "ימים בהשאלה", "תיאור כלי", "שם הכלי")

        self.borrowed_tools_tree = ttk.Treeview(self.borrowed_tools_frame, columns=borrowed_tools_tree_columns_logical, show="headings")
        borrowed_tools_tree_style = ttk.Style()
        borrowed_tools_tree_style.configure("Treeview.Heading", font=('Helvetica', 12, 'bold'))
        borrowed_tools_tree_style.configure("Treeview", font=('Helvetica', 12), rowheight=25)

        self.borrowed_tools_tree.heading("שם הכלי", text="שם הכלי", anchor="e")
        self.borrowed_tools_tree.column("שם הכלי", width=150, anchor="e")
        self.borrowed_tools_tree.heading("תיאור כלי", text="תיאור כלי", anchor="e")
        self.borrowed_tools_tree.column("תיאור כלי", width=200, anchor="e")
        self.borrowed_tools_tree.heading("ימים בהשאלה", text="ימים בהשאלה", anchor="e")
        self.borrowed_tools_tree.column("ימים בהשאלה", width=100, anchor="e")
        self.borrowed_tools_tree.heading("תאריך השאלה", text="תאריך השאלה", anchor="e")
        self.borrowed_tools_tree.column("תאריך השאלה", width=150, anchor="e")

        borrowed_tools_tree_scrollbar = ttkb.Scrollbar(self.borrowed_tools_frame, orient="vertical", command=self.borrowed_tools_tree.yview)
        self.borrowed_tools_tree.configure(yscrollcommand=borrowed_tools_tree_scrollbar.set)

        borrowed_tools_tree_scrollbar.grid(row=1, column=1, sticky="ns")
        self.borrowed_tools_tree.grid(row=1, column=0, sticky="nsew")

        self.borrowed_tools_frame.grid_columnconfigure(0, weight=1)
        self.borrowed_tools_frame.grid_rowconfigure(1, weight=1)


    # --- מתודות של המחלקה ToolBorrowingApp ---

    def show_tools_context_menu(self, event):
        """מציג תפריט קליק ימני עבור טבלת הכלים."""
        item_id = self.tools_tree.identify_row(event.y)

        if item_id:
            self.tools_tree.selection_set(item_id)
            context_menu = tk.Menu(self.root, tearoff=0)
            item_values = self.tools_tree.item(item_id, 'values')

            # אינדקסים לפי העמודות המוצגות כעת בטבלה הראשית (6 עמודות)
            # ימים(0), שם השואל(1), סטטוס(2), תיאור כלי(3), שם הכלי(4), מספר סידורי(5)

            if item_values and len(item_values) >= 6:
                 tool_status = self.tree_item_value_safe(item_values, 2)

                 if tool_status == "מושאל":
                     context_menu.add_command(label="החזר כלי", command=self.return_tool)
                 elif tool_status == "זמין":
                     context_menu.add_command(label="השאל כלי", command=self.borrow_tool)

                 try:
                     context_menu.tk_popup(event.x_root, event.y_root)
                 finally:
                     context_menu.grab_release()
            else:
                 pass


    def refresh_ui(self):
        """מרענן את כל הטבלאות ושדות הקלט בממשק."""
        self.filter_tools(None)
        self.filter_borrower_table(None)
        self.on_borrower_table_select(None)


    def update_after_borrow_return(self):
        """מרענן את ממשק המשתמש לאחר פעולת השאלה או החזרה."""
        # 1. רענן את טבלת הכלים הראשית כדי להציג סטטוס ושואל מעודכנים.
        self.filter_tools(None)

        # 2. קבל את שם השואל שנבחר כרגע בטבלת השואלים.
        selected_item_id = self.borrower_tree.selection()
        selected_borrower_name = ""
        if selected_item_id:
             # קבל באופן בטוח את שם השואל מהפריט הנבחר (אינדקס 2 בטבלת השואלים)
             item_values = self.borrower_tree.item(selected_item_id[0], 'values')
             if item_values and len(item_values) > 2:
                  selected_borrower_name = self.tree_item_value_safe(item_values, 2)

        # 3. רענן את טבלת הכלים המושאלים על בסיס שם השואל שנבחר.
        if selected_borrower_name:
             # ודא שהשואל עדיין קיים בנתונים (תרחיש פחות סביר לאחר השאלה/החזרה)
             selected_borrower_object = self.data_manager.find_borrower(selected_borrower_name)
             if selected_borrower_object:
                  # מצא את הכלים המושאלים לשואל זה מהנתונים הפנימיים
                  borrowed_list = [tool for tool in self.data_manager.tools_data if tool.get("שם השואל") == selected_borrower_name and tool.get("סטטוס") == "מושאל"]
                  # רענן את טבלת הכלים המושאלים עם הרשימה המעודכנת
                  self.refresh_borrowed_tools_table(borrowed_list)
                  # חשוב: בחירת השואל בטבלת השואלים נשמרת כי לא קראנו ל-filter_borrower_table שמנקה אותה.
                  # גם שדות הקלט של השואל יישארו מלאים.
             else:
                  # השואל לא נמצא (אולי נמחק בינתיים), נקה את טבלת הכלים המושאלים ואת הבחירה.
                  self.refresh_borrowed_tools_table([])
                  self.borrower_tree.selection_set([]) # בטל את הבחירה בטבלת השואלים
                  self.on_borrower_table_select(None) # נקה את שדות הקלט של השואל
        else:
             # לא נבחר שואל לפני הפעולה, ודא שטבלת הכלים המושאלים ריקה.
             self.refresh_borrowed_tools_table([])
             # במקרה זה, שדות הקלט של השואל וטבלת השואלים אמורים כבר להיות ריקים/ללא בחירה.
             # אין צורך לקרוא ל-on_borrower_table_select(None) שוב.

        # הערה חשובה: אין לקרוא כאן ל-self.filter_borrower_table(None)
        # כדי למנוע איבוד בחירת השואל בטבלת השואלים.



    def undo_last_action(self):
        """מטפל בלחיצה על כפתור ביטול פעולות ומבקש אישור לפני ביצוע."""
        if not self.data_manager.history:
             self.show_toast("אין למה לבטל (לא נטען מצב התחלתי).", title="ביטול פעולה", bootstyle="info")
             return

        confirm = messagebox.askyesno(
            "אישור ביטול פעולות",
            "?האם אתה בטוח שברצונך לבטל את כל הפעולות\nפעולה זו תחזיר את הנתונים למצבם ההתחלתי"
        )

        if confirm:
            if self.data_manager.restore_initial_state():
                self.show_toast("!כל הפעולות שבוצעו - בוטלו", title="ביטול פעולות", bootstyle="danger", position=(500, 250, 'nw'))
                self.refresh_ui()
                self.tools_tree.selection_set([])
                self.borrower_tree.selection_set([])
        else:
            self.show_toast("ביטול פעולות בוטל על ידי המשתמש.", title="פעולה בוטלה", bootstyle="info")

    def refresh_tools_table(self, data_to_display):
        """מרענן את טבלת הכלים הראשית עם הנתונים שסופקו."""
        for item in self.tools_tree.get_children():
            self.tools_tree.delete(item)

        sorted_data_to_display = sorted(data_to_display, key=lambda tool: (0 if tool.get("סטטוס") == "מושאל" else 1, tool.get("שם הכלי", ""), tool.get("תיאור כלי", "")))

        for tool in sorted_data_to_display:
            days_borrowed = ""
            if tool.get("סטטוס") == "מושאל":
                 days_borrowed = self.data_manager.calculate_days_borrowed(tool.get("תאריך השאלה", ""))

            self.tools_tree.insert("", tk.END, values=(
                days_borrowed,
                tool.get("שם השואל", ""),
                tool.get("סטטוס", ""),
                tool.get("תיאור כלי", ""),
                tool.get("שם הכלי", ""),
                tool.get("מספר סידורי", "")
            ))
            if tool.get("סטטוס") == "מושאל":
                self.tools_tree.item(self.tools_tree.get_children()[-1], tags=('borrowed',))


    def populate_borrower_table(self, data):
        """ממלא את טבלת השואלים בנתונים."""
        for item in self.borrower_tree.get_children():
            self.borrower_tree.delete(item)

        for borrower in data:
            self.borrower_tree.insert("", tk.END, values=(
                borrower.get("כתובת", ""),
                borrower.get("מספר טלפון", ""),
                borrower.get("שם השואל", "")
            ))


    def refresh_borrowed_tools_table(self, borrowed_list):
        """מרענן את טבלת הכלים הממושאלים לשואל הנבחר."""
        for item in self.borrowed_tools_tree.get_children():
            self.borrowed_tools_tree.delete(item)

        sorted_borrowed_list = sorted(borrowed_list, key=lambda tool: tool.get("תאריך השאלה", ""), reverse=True)

        for tool in sorted_borrowed_list:
            days_borrowed = self.data_manager.calculate_days_borrowed(tool.get("תאריך השאלה", ""))
            self.borrowed_tools_tree.insert("", tk.END, values=(
                tool.get("תאריך השאלה", ""),
                days_borrowed,
                tool.get("תיאור כלי", ""),
                tool.get("שם הכלי", "")
            ))

    def filter_tools(self, event=None):
        """מסנן את רשימת הכלים לפי מונח חיפוש בשדות שם, תיאור ומספר סידורי."""
        search_term = self.search_entry.get().strip().lower()
        filtered_data = []
        if not search_term:
            filtered_data = self.data_manager.tools_data
        else:
            for tool in self.data_manager.tools_data:
                tool_name = tool.get("שם הכלי", "").lower()
                tool_description = tool.get("תיאור כלי", "").lower()
                serial_number = tool.get("מספר סידורי", "").lower()

                if (search_term in tool_name or
                    search_term in tool_description or
                    search_term in serial_number):
                    filtered_data.append(tool)
        self.refresh_tools_table(filtered_data)


    def filter_borrower_table(self, event=None):
        """מסנן את טבלת השואלים לפי שדות הקלט."""
        search_name = self.borrower_name_entry.get().strip().lower()
        search_phone = self.borrower_phone_entry.get().strip().lower()
        search_address = self.borrower_address_entry.get().strip().lower()

        filtered_data = []
        for borrower in self.data_manager.borrowers_data:
            name = borrower.get("שם השואל", "").lower()
            phone = borrower.get("מספר טלפון", "").lower()
            address = borrower.get("כתובת", "").lower()

            name_match = not search_name or search_name in name
            phone_match = not search_phone or search_phone in phone
            address_match = not search_address or search_address in address

            if name_match and phone_match and address_match:
                filtered_data.append(borrower)
        self.populate_borrower_table(filtered_data)

    def on_borrower_table_select(self, event):
        """מעדכן פרטי שואל וטבלת כלים מושאלים בעת בחירה."""
        selected_item = self.borrower_tree.selection()
        if not selected_item:
            self.borrower_name_entry.delete(0, tk.END)
            self.borrower_phone_entry.delete(0, tk.END)
            self.borrower_address_entry.delete(0, tk.END)
            self.refresh_borrowed_tools_table([])
            return

        selected_borrower_name = ""
        # אינדקס 2 הוא שם השואל בטבלת השואלים הראשית
        if selected_item and self.borrower_tree.item(selected_item[0], 'values') and len(self.borrower_tree.item(selected_item[0], 'values')) > 2:
            selected_borrower_name = self.tree_item_value_safe(self.borrower_tree.item(selected_item[0], 'values'), 2)
        else:
             self.borrower_name_entry.delete(0, tk.END)
             self.borrower_phone_entry.delete(0, tk.END)
             self.borrower_address_entry.delete(0, tk.END)
             self.refresh_borrowed_tools_table([])
             return


        selected_borrower = self.data_manager.find_borrower(selected_borrower_name)

        if selected_borrower:
            self.borrower_name_entry.delete(0, tk.END)
            self.borrower_name_entry.insert(0, selected_borrower.get("שם השואל", ""))
            self.borrower_phone_entry.delete(0, tk.END)
            self.borrower_phone_entry.insert(0, selected_borrower.get("מספר טלפון", ""))
            self.borrower_address_entry.delete(0, tk.END)
            self.borrower_address_entry.insert(0, selected_borrower.get("כתובת", ""))

            borrowed_list = [tool for tool in self.data_manager.tools_data if tool.get("שם השואל") == selected_borrower_name and tool.get("סטטוס") == "מושאל"]
            self.refresh_borrowed_tools_table(borrowed_list)
        else:
            self.borrower_name_entry.delete(0, tk.END)
            self.borrower_phone_entry.delete(0, tk.END)
            self.borrower_address_entry.delete(0, tk.END)
            self.refresh_borrowed_tools_table([])


    def add_borrower(self):
        """מוסיף שואל חדש."""
        borrower_name = self.borrower_name_entry.get().strip()
        phone = self.borrower_phone_entry.get().strip()
        address = self.borrower_address_entry.get().strip()

        if not borrower_name:
            self.show_toast("אנא הזן שם לשואל.", title="קלט חסר", bootstyle="warning", position=(400, 405, 'nw'))
            return

        existing_borrower = self.data_manager.find_borrower(borrower_name)

        if existing_borrower:
            self.show_toast(f"השואל '{borrower_name}' כבר קיים ברשימה.", title="שואל קיים", bootstyle="info", position=(400, 405, 'nw'))
            self.filter_borrower_table(None)
        else:
            new_borrower = {
                "שם השואל": borrower_name,
                "מספר טלפון": phone,
                "כתובת": address
            }
            self.data_manager.borrowers_data.append(new_borrower)
            self.show_toast(f"השואל '{borrower_name}' נוסף בהצלחה.", title="הוספת שואל", bootstyle="success", position=(400, 405, 'nw'))
            self.filter_borrower_table(None)
            self.data_manager.save_all_data()

    def borrow_tool(self):
        """משאיל כלים נבחרים לשואל נבחר."""
        selected_tool_items = self.tools_tree.selection()
        selected_borrower_item = self.borrower_tree.selection()

        if not selected_tool_items:
            self.show_toast("אנא בחר כלי אחד או יותר מהטבלה הראשית להשאלה.", title="בחירה חסרה", bootstyle="warning", position=(400, 405, 'nw'))
            return

        if not selected_borrower_item:
            self.show_toast("אנא בחר שואל מהטבלה התחתונה.", title="בחירה חסרה", bootstyle="warning", position=(400, 405, 'nw'))
            return

        selected_borrower_name = ""
        # אינדקס 2 הוא שם השואל בטבלת השואלים הראשית
        if selected_borrower_item and self.borrower_tree.item(selected_borrower_item[0], 'values') and len(self.borrower_tree.item(selected_borrower_item[0], 'values')) > 2:
            selected_borrower_name = self.tree_item_value_safe(self.borrower_tree.item(selected_borrower_item[0], 'values'), 2)
        else:
            self.show_toast("לא ניתן היה לזהות את השואל הנבחר.", title="שגיאה פנימית", bootstyle="danger")
            return

        if not self.data_manager.find_borrower(selected_borrower_name):
            self.show_toast(f"השואל '{selected_borrower_name}' לא נמצא במערכת.", title="שגיאה", bootstyle="danger")
            self.refresh_ui()
            return

        borrowed_count = 0
        already_borrowed_tools = []
        borrow_date = date.today().strftime('%Y-%m-%d')

        for item_id in selected_tool_items:
            item_values = self.tools_tree.item(item_id, 'values')
            # אינדקסים לפי העמודות המוצגות בטבלה הראשית
            # ימים(0), שם השואל(1), סטטוס(2), תיאור כלי(3), שם הכלי(4), מספר סידורי(5)
            if item_values and len(item_values) >= 6:
                tool_name = self.tree_item_value_safe(item_values, 4)
                tool_description = self.tree_item_value_safe(item_values, 3)
                tool_serial = self.tree_item_value_safe(item_values, 5)
            else:
                 self.show_toast(f"שגיאה: לא ניתן היה לקרוא פרטי כלי (אינדקסים לא מתאימים).", title="שגיאה פנימית", bootstyle="danger")
                 continue

            tool_to_borrow = self.data_manager.find_tool_by_serial(tool_serial)
            if not tool_to_borrow:
                 tool_to_borrow = self.data_manager.find_tool(tool_name, tool_description)


            if tool_to_borrow:
                if tool_to_borrow.get("סטטוס") == "זמין":
                    tool_to_borrow["סטטוס"] = "מושאל"
                    tool_to_borrow["שם השואל"] = selected_borrower_name
                    tool_to_borrow["תאריך השאלה"] = borrow_date
                    try:
                         current_count = int(tool_to_borrow.get("מונה השאלות", "0"))
                         tool_to_borrow["מונה השאלות"] = str(current_count + 1)
                    except (ValueError, TypeError):
                        tool_to_borrow["מונה השאלות"] = "1"
                    borrowed_count += 1
                    self.data_manager.append_borrowing_history(tool_name, selected_borrower_name, borrow_date)
                else:
                    already_borrowed_tools.append(f"{tool_name} ({tool_description})")
            else:
                self.show_toast(f"הכלי '{tool_name} ({tool_description})' לא נמצא בנתונים הפנימיים ולא ניתן להשאיל אותו.", title="שגיאה פנימית", bootstyle="warning")

        if borrowed_count > 0:
            self.data_manager.save_all_data()
            self.update_after_borrow_return()
            self.tools_tree.selection_set([])
            self.show_toast(f"הושאלו בהצלחה {borrowed_count} כלים לשואל {selected_borrower_name}.", title="השאלת כלי", bootstyle="primary", position=(400, 405, 'nw'))

        if already_borrowed_tools:
            self.show_toast(f"הכלים הבאים כבר מושאלים ועל כן לא הושאלו מחדש: {', '.join(already_borrowed_tools)}", title="שגיאת השאלה", bootstyle="warning", position=(400, 405, 'nw'))

        if borrowed_count == 0 and not already_borrowed_tools:
             self.show_toast("לא נבחרו כלים זמינים להשאלה או שלא ניתן היה להשאיל אותם.", title="שגיאה", bootstyle="warning", position=(400, 405, 'nw'))


    def return_tool(self):
        """מחזיר כלים נבחרים."""
        selected_items_main = self.tools_tree.selection()
        selected_items_borrowed = self.borrowed_tools_tree.selection()

        if not selected_items_main and not selected_items_borrowed:
            self.show_toast("אנא בחר כלי אחד או יותר להחזרה מהטבלה הראשית או מהטבלה התחתונה.", title="בחירה חסרה", bootstyle="warning", position=(400, 405, 'nw'))
            return

        selected_borrower_item = self.borrower_tree.selection()
        current_borrower_name = ""
        # אינדקס 2 הוא שם השואל בטבלת השואלים הראשית
        if selected_borrower_item and self.borrower_tree.item(selected_borrower_item[0], 'values') and len(self.borrower_tree.item(selected_borrower_item[0], 'values')) > 2:
             current_borrower_name = self.tree_item_value_safe(self.borrower_tree.item(selected_borrower_item[0], 'values'), 2)

        returned_count = 0
        unique_tools_to_return = set()

        for item_id in selected_items_main:
            item_values = self.tools_tree.item(item_id, 'values')
             # אינדקסים לפי העמודות המוצגות בטבלה הראשית
            # ימים(0), שם השואל(1), סטטוס(2), תיאור כלי(3), שם הכלי(4), מספר סידורי(5)
            if item_values and len(item_values) >= 6:
                tool_name = self.tree_item_value_safe(item_values, 4)
                tool_description = self.tree_item_value_safe(item_values, 3)
                tool_serial = self.tree_item_value_safe(item_values, 5)
                tool_status_display = self.tree_item_value_safe(item_values, 2)
                borrower_name_display = self.tree_item_value_safe(item_values, 1)

                if tool_status_display == "מושאל" and borrower_name_display != "":
                    if tool_serial:
                         unique_tools_to_return.add((tool_serial, "", ""))
                    else:
                         unique_tools_to_return.add(("", tool_name, tool_description))
                else:
                     pass # כלי זמין או ללא שואל משויך - אין צורך לנסות להחזיר דרך הטבלה הראשית
            else:
                pass # שגיאה בקריאת שורת הטבלה

        # טיפול בכלים שנבחרו בטבלת הכלים המושאלים
        if current_borrower_name:
            for item_id in selected_items_borrowed:
                item_values = self.borrowed_tools_tree.item(item_id, 'values')
                # אינדקסים לפי העמודות בטבלת הכלים המושאלים
                # תאריך השאלה(0), ימים בהשאלה(1), תיאור כלי(2), שם הכלי(3)
                if item_values and len(item_values) >= 4:
                    tool_name_borrowed = self.tree_item_value_safe(item_values, 3)
                    tool_description_borrowed = self.tree_item_value_safe(item_values, 2)
                    # איתור הכלי המלא בנתונים הפנימיים לפי שם ותיאור (בלבד בטבלת מושאלים)
                    matching_tool_in_data = next(
                        (tool for tool in self.data_manager.tools_data
                         if tool.get('שם הכלי') == tool_name_borrowed
                         and tool.get('תיאור כלי', '') == tool_description_borrowed
                         and tool.get('סטטוס') == 'מושאל' # לוודא שהכלי מושאל (תיאורטית יכול להיות כלי עם אותו שם ותיאור זמין)
                         and tool.get('שם השואל') == current_borrower_name), # לוודא שהוא מושאל לשואל הנבחר
                        None
                    )
                    if matching_tool_in_data:
                        serial_borrowed = matching_tool_in_data.get("מספר סידורי", "")
                        if serial_borrowed:
                             unique_tools_to_return.add((serial_borrowed, "", ""))
                        else:
                             unique_tools_to_return.add(("", tool_name_borrowed, tool_description_borrowed))
                    else:
                         pass # הכלי לא נמצא בנתונים הפנימיים תואם לשואל הנבחר / סטטוס מושאל
                else:
                    pass # שגיאה בקריאת שורת הטבלה

        tools_already_available = []
        tools_not_found_or_unexpected_status = []

        for item_identifier in unique_tools_to_return:
             serial, name, description = item_identifier
             tool_to_return = None
             if serial:
                  tool_to_return = self.data_manager.find_tool_by_serial(serial)
             # אם לא נמצא לפי מספר סידורי, נסה לפי שם ותיאור (אם השם קיים - התיאור עשוי להיות ריק)
             if not tool_to_return and name:
                  # כדי למצוא כלי לפי שם ותיאור, שניהם חייבים להיות מוגדרים בנתונים
                  tool_to_return = self.data_manager.find_tool(name, description)


             if tool_to_return:
                 tool_id_str = f"{tool_to_return.get('שם הכלי', '')} ({tool_to_return.get('תיאור כלי', '')})"
                 if tool_to_return.get('מספר סידורי'):
                      tool_id_str += f" [{tool_to_return.get('מספר סידורי')}]"

                 if tool_to_return.get("סטטוס") == "מושאל":
                     # לוודא שהכלי אכן מושאל
                     if tool_to_return.get("שם השואל") != "":
                         tool_to_return["סטטוס"] = "זמין"
                         tool_to_return["שם השואל"] = ""
                         tool_to_return["תאריך השאלה"] = ""
                         returned_count += 1
                     else:
                         # כלי מושאל בנתונים הפנימיים אבל ללא שואל משויך - מצב לא תקין
                         tools_not_found_or_unexpected_status.append(f"{tool_id_str} - מושאל ללא שואל משויך")
                 elif tool_to_return.get("סטטוס") == "זמין":
                      # הכלי כבר זמין - אין צורך להחזיר
                      tools_already_available.append(f"{tool_id_str}")
                 else:
                      # סטטוס אחר כלשהו
                      tools_not_found_or_unexpected_status.append(f"{tool_id_str} - סטטוס לא צפוי: {tool_to_return.get('סטטוס')}")
             else:
                  # הכלי לא נמצא בנתונים הפנימיים בכלל
                  item_str_parts = []
                  if serial: item_str_parts.append(f"מספר סידורי: {serial}")
                  if name: item_str_parts.append(f"שם: {name}")
                  if description is not None: item_str_parts.append(f"תיאור: {description}") # תיאור יכול להיות מחרוזת ריקה
                  item_str = ", ".join(item_str_parts) if item_str_parts else "לא ידוע"
                  tools_not_found_or_unexpected_status.append(f"פריט לא נמצא: ({item_str})")

        if returned_count > 0:
            self.data_manager.save_all_data()
            # רענן את טבלת הכלים הראשית כדי לשקף את הכלים שהוחזרו
            self.filter_tools(None)
            # כעת קרא למתודה החדשה לרענון טבלת הכלים המושאלים עבור השואל הנבחר
            self.update_after_borrow_return()  # זו תרענן את טבלת המושאלים עבור השואל הנבחר
            self.tools_tree.selection_set([])  # בטל בחירה בטבלת הכלים הראשית - השאר זאת
            self.borrowed_tools_tree.selection_set([])  # בטל בחירה בטבלת הכלים המושאלים - השאר זאת
            self.show_toast(f"הוחזרו בהצלחה {returned_count} כלים לגמ\"ח.", title="החזרת כלי", bootstyle="success",
                            position=(400, 405, 'nw'))

        if tools_already_available:
            self.show_toast(f"הכלים הבאים כבר היו זמינים ועל כן לא נדרשה החזרה: {', '.join(tools_already_available)}", title="אזהרת החזרה", bootstyle="warning", position=(400, 405, 'nw'))

        if tools_not_found_or_unexpected_status:
            self.show_toast(f"לא ניתן היה למצוא או להחזיר את הכלים הבאים (יתכן שלא נמצאו או במצב לא צפוי): {', '.join(tools_not_found_or_unexpected_status)}", title="אזהרת החזרה", bootstyle="warning")

        if returned_count == 0 and not tools_already_available and not tools_not_found_or_unexpected_status:
             self.show_toast("לא נבחרו כלים מושאלים להחזרה או שלא ניתן היה להחזירם (יתכן שהם כבר זמינים).", title="שגיאת החזרה", bootstyle="warning", position=(400, 405, 'nw'))



    def tree_item_value_safe(self, values, index):
        """מחזירה ערך מתוך טאפל של ערכי Treeview באינדקס נתון, מחזירה מחרוזת ריקה אם האינדקס לא קיים."""
        if isinstance(values, (list, tuple)) and index < len(values):
            return str(values[index]) if values[index] is not None else ""
        return ""

    def show_statistics(self):
        """מציג סטטיסטיקת השאלות בחלון נפרד."""
        stats_window = ttkb.Toplevel(self.root)
        stats_window.title("סטטיסטיקת השאלות")
        stats_window.attributes('-topmost', True)

        window_width = 250
        window_height = 510
        screen_width = stats_window.winfo_screenwidth()
        screen_height = stats_window.winfo_screenheight()
        center_x = int(screen_width / 2 - window_width / 2)
        center_y = int(screen_height / 2 - window_height / 2)
        stats_window.geometry(f"{window_width}x{window_height}+{center_x}+{center_y}")

        main_frame = ttkb.Frame(stats_window, padding="20")
        main_frame.pack(expand=True, fill="both")

        summary_frame = ttkb.Frame(main_frame)
        summary_frame.grid_columnconfigure(0, weight=1)
        summary_frame.grid_columnconfigure(1, weight=0)
        summary_frame.grid_columnconfigure(2, weight=0)
        summary_frame.grid_columnconfigure(3, weight=0)
        summary_frame.grid_columnconfigure(4, weight=1)
        summary_frame.pack(fill="x", pady=(0, 15))

        tool_borrow_counts_by_name = {}
        total_borrowed_tools_current = 0

        for tool in self.data_manager.tools_data:
            tool_name = tool.get("שם הכלי", "כלי ללא שם")
            try:
                borrow_count = int(tool.get("מונה השאלות", "0"))
            except (ValueError, TypeError):
                borrow_count = 0
            tool_borrow_counts_by_name[tool_name] = tool_borrow_counts_by_name.get(tool_name, 0) + borrow_count
            if tool.get("סטטוס") == "מושאל":
                total_borrowed_tools_current += 1

        total_borrowed_tools_historical = len(self.data_manager.borrowing_history_data)

        current_borrowed_label = ttkb.Label(summary_frame, text="בהשאלה", font=('Helvetica', 10), anchor="center")
        current_borrowed_label.grid(row=1, column=1, sticky="nsew", padx=(0, 5))
        current_borrowed_value = ttkb.Label(summary_frame, text=str(total_borrowed_tools_current),
                                            font=('Helvetica', 14, 'bold'), bootstyle="danger", anchor="center")
        current_borrowed_value.grid(row=0, column=1, sticky="nsew")

        spacer_label = ttkb.Label(summary_frame, text=" | ", font=('Helvetica', 18), bootstyle="secondary")
        spacer_label.grid(row=0, column=2, rowspan=2, sticky="ns", padx=10)

        total_historical_label = ttkb.Label(summary_frame, text="השאלות", font=('Helvetica', 10), anchor="center")
        total_historical_label.grid(row=1, column=3, sticky="nsew", padx=(0, 5))
        total_historical_value = ttkb.Label(summary_frame, text=str(total_borrowed_tools_historical),
                                            font=('Helvetica', 14, 'bold'), bootstyle="info", anchor="center")
        total_historical_value.grid(row=0, column=3, sticky="nsew")

        stats_list = [(name, count) for name, count in tool_borrow_counts_by_name.items()]
        stats_list.sort(key=lambda item: item[1], reverse=True)

        tool_stats_table_frame = ttkb.Frame(main_frame)
        tool_stats_table_frame.pack(expand=True, fill="both")
        tool_stats_table_frame.columnconfigure(0, weight=1)
        tool_stats_table_frame.rowconfigure(0, weight=1)

        tree_columns_logical = ("מספר השאלות", "שם הכלי")
        tool_stats_tree = ttk.Treeview(tool_stats_table_frame, columns=tree_columns_logical, show="")
        style = ttk.Style()
        style.configure("Treeview.Heading", font=('Helvetica', 12, 'bold'))
        style.configure("Treeview", font=('Helvetica', 12), rowheight=25)
        tool_stats_tree.tag_configure("evenrow", background="#f8f9fa")
        tool_stats_tree.tag_configure("oddrow", background="#e9ecef")

        tool_stats_tree.heading("שם הכלי", text="שם כלי", anchor="e")
        tool_stats_tree.column("שם הכלי", width=80, anchor="e")
        tool_stats_tree.heading("מספר השאלות", text="הושאל פעמים", anchor="e")
        tool_stats_tree.column("מספר השאלות", width=30, anchor="e")

        for i, (name, count) in enumerate(stats_list):
            tag = 'evenrow' if i % 2 == 0 else 'oddrow'
            tool_stats_tree.insert("", tk.END, values=(count, name), tags=(tag,))

        scrollbar = ttkb.Scrollbar(tool_stats_table_frame, orient="vertical", command=tool_stats_tree.yview)
        tool_stats_tree.configure(yscrollcommand=scrollbar.set)
        scrollbar.grid(row=0, column=1, sticky="ns")
        tool_stats_tree.grid(row=0, column=0, sticky="nsew")

        if not stats_list and total_borrowed_tools_current == 0 and total_borrowed_tools_historical == 0:
            no_data_label = ttkb.Label(main_frame, text="אין נתוני השאלות להצגה.", font=('Helvetica', 12),
                                       bootstyle="info")
            no_data_label.pack(pady=10)
        stats_window.transient(self.root)


    def show_borrowing_history(self):
        """פותח חלון חדש להצגת היסטוריית השאלות."""
        history_window = ttkb.Toplevel(self.root)
        history_window.title("היסטוריית השאלות")

        # --- הוספת קוד למיקום החלון במרכז המסך ---
        window_width = 800  # רוחב החלון (ניתן להתאים)
        window_height = 600 # גובה החלון (ניתן להתאים)
        screen_width = history_window.winfo_screenwidth()
        screen_height = history_window.winfo_screenheight()
        center_x = int(screen_width / 2 - window_width / 2)
        center_y = int(screen_height / 2 - window_height / 2)
        history_window.geometry(f"{window_width}x{window_height}+{center_x}+{center_y}")
        # --------------------------------------------------

        history_search_frame = ttkb.Frame(history_window, padding="10")
        history_search_frame.grid(row=0, column=0, sticky="ew", padx=10, pady=(10, 0))
        history_window.grid_columnconfigure(0, weight=1)

        history_search_label = ttkb.Label(history_search_frame, text=":חיפוש בהיסטוריה", font=('Helvetica', 12))
        history_search_label.grid(row=0, column=2, sticky="e", padx=(0, 5))

        history_search_entry = ttkb.Entry(history_search_frame, width=50, justify='right')
        history_search_entry.grid(row=0, column=1, sticky="ew", padx=5)
        history_search_entry.bind("<Escape>", self.clear_entry_on_escape)

        history_search_frame.grid_columnconfigure(1, weight=1)

        history_table_frame = ttkb.Frame(history_window, padding="10")
        history_table_frame.grid(row=1, column=0, sticky="nsew", padx=10, pady=(0, 10))
        history_window.grid_rowconfigure(1, weight=1)

        history_tree_columns_logical = BORROWING_HISTORY_FIELDNAMES

        history_tree = ttk.Treeview(history_table_frame, columns=history_tree_columns_logical, show="headings")
        history_tree_style = ttk.Style()
        history_tree_style.configure("Treeview.Heading", font=('Helvetica', 12, 'bold'))
        history_tree_style.configure("Treeview", font=('Helvetica', 12), rowheight=25)

        history_tree.heading("תאריך השאלה", text="תאריך השאלה", anchor="e")
        history_tree.column("תאריך השאלה", width=150, anchor="e")
        history_tree.heading("שם הכלי", text="שם הכלי", anchor="e")
        history_tree.column("שם הכלי", width=200, anchor="e")
        history_tree.heading("שם השואל", text="שם השואל", anchor="e")
        history_tree.column("שם השואל", width=200, anchor="e")

        history_tree_scrollbar = ttkb.Scrollbar(history_table_frame, orient="vertical", command=history_tree.yview)
        history_tree.configure(yscrollcommand=history_tree_scrollbar.set)

        history_tree_scrollbar.grid(row=0, column=1, sticky="ns")
        history_tree.grid(row=0, column=0, sticky="nsew")

        history_table_frame.grid_columnconfigure(0, weight=1)
        history_table_frame.grid_rowconfigure(0, weight=1)

        history_search_entry.bind("<KeyRelease>", lambda event: self.filter_history_table(history_tree, history_search_entry.get()))

        reversed_history_data = self.data_manager.borrowing_history_data[::-1]
        self.populate_history_table(history_tree, reversed_history_data)

        history_window.transient(self.root)


    def populate_history_table(self, tree, data):
        """ממלא את טבלת היסטוריית ההשאלות בנתונים."""
        for item in tree.get_children():
            tree.delete(item)
        for record in data:
            tree.insert("", tk.END, values=(
                record.get("תאריך השאלה", ""),
                record.get("שם הכלי", ""),
                record.get("שם השואל", "")
            ))

    def filter_history_table(self, tree, search_term):
        """מסנן את טבלת ההיסטוריה לפי מונח חיפוש ומסדר לפי תאריך יורד."""
        filtered_data = []
        if not search_term:
            filtered_data = self.data_manager.borrowing_history_data
        else:
            lower_search_term = search_term.lower()
            for record in self.data_manager.borrowing_history_data:
                date_borrowed = record.get("תאריך השאלה", "").lower()
                tool_name = record.get("שם הכלי", "").lower()
                borrower_name = record.get("שם השואל", "").lower()
                if lower_search_term in date_borrowed or lower_search_term in tool_name or lower_search_term in borrower_name:
                    filtered_data.append(record)

        reversed_filtered_data = filtered_data[::-1]
        self.populate_history_table(tree, reversed_filtered_data)


    def show_add_tool_window(self):
        """פותח חלון חדש לניהול כלים."""
        # בדיקה אם החלון כבר פתוח
        if hasattr(self, 'tool_management_window') and self.tool_management_window is not None and self.tool_management_window.winfo_exists():
            self.tool_management_window.lift() # הבאת החלון לפורגראונד
            return

        tool_window = ttkb.Toplevel(self.root)
        tool_window.title("ניהול כלים")

        # --- הוספת קוד למיקום החלון במרכז המסך ---
        window_width = 750  # רוחב החלון (ניתן להתאים)
        window_height = 600 # גובה החלון (ניתן להתאים)
        screen_width = tool_window.winfo_screenwidth()
        screen_height = tool_window.winfo_screenheight()
        center_x = int(screen_width / 2 - window_width / 2)
        center_y = int(screen_height / 2 - window_height / 2)
        tool_window.geometry(f"{window_width}x{window_height}+{center_x}+{center_y}")
        # --------------------------------------------------

        # שמירת רפרנס לחלון הניהול
        self.tool_management_window = tool_window

        tool_controls_frame = ttkb.Frame(tool_window, padding="15")
        tool_controls_frame.grid(row=0, column=0, sticky="ew", padx=10, pady=5)
        tool_window.grid_columnconfigure(0, weight=1)

        # --- פריסת grid ישירה בתוך tool_controls_frame (סדר הפוך) ---
        # הגדרת עמודות: כפתור (0), שדה קלט (1), רווח (2), תווית (3)
        tool_controls_frame.grid_columnconfigure(0, weight=0) # כפתור - לא תתרחב
        tool_controls_frame.grid_columnconfigure(1, weight=1) # שדה קלט - יתרחב
        tool_controls_frame.grid_columnconfigure(2, weight=0) # רווח קטן - לא יתרחב (אפשר גם להוריד אם לא נחוץ)
        tool_controls_frame.grid_columnconfigure(3, weight=0) # תווית - לא תתרחב


        # שורה 0: הוסף כלי + שם כלי
        add_update_button_window = ttkb.Button(tool_controls_frame, text="הוסף", bootstyle="success")
        add_update_button_window.grid(row=0, column=0, sticky="w", padx=5) # מיקום בעמודה 0

        tool_name_entry_window = ttkb.Entry(tool_controls_frame, justify='right')
        tool_name_entry_window.grid(row=0, column=1, sticky="ew", padx=5, pady=2) # מיקום בעמודה 1
        tool_name_entry_window.bind("<Escape>", self.clear_entry_on_escape)

        tool_name_label_window = ttkb.Label(tool_controls_frame, text=":שם כלי", anchor="e")
        tool_name_label_window.grid(row=0, column=3, sticky="e", padx=(10, 2), pady=2) # מיקום בעמודה 3


        # שורה 1: ערוך כלי + תיאור כלי
        edit_tool_button = ttkb.Button(tool_controls_frame, text="ערוך", bootstyle="info", state=tk.DISABLED)
        edit_tool_button.grid(row=1, column=0, sticky="w", padx=5) # מיקום בעמודה 0

        tool_description_entry_window = ttkb.Entry(tool_controls_frame, justify='right')
        tool_description_entry_window.grid(row=1, column=1, sticky="ew", padx=5, pady=2) # מיקום בעמודה 1
        tool_description_entry_window.bind("<Escape>", self.clear_entry_on_escape)

        tool_description_label_window = ttkb.Label(tool_controls_frame, text=":תיאור כלי", anchor="e")
        tool_description_label_window.grid(row=1, column=3, sticky="e", padx=(10, 2), pady=2) # מיקום בעמודה 3


        # שורה 2: מחק כלי + מספר סידורי
        delete_tool_button = ttkb.Button(tool_controls_frame, text="מחק", bootstyle="danger", state=tk.DISABLED)
        delete_tool_button.grid(row=2, column=0, sticky="w", padx=5) # מיקום בעמודה 0

        serial_number_entry_window = ttkb.Entry(tool_controls_frame, justify='right')
        serial_number_entry_window.grid(row=2, column=1, sticky="ew", padx=5, pady=2) # מיקום בעמודה 1
        serial_number_entry_window.bind("<Escape>", self.clear_entry_on_escape)

        serial_number_label_window = ttkb.Label(tool_controls_frame, text=":מספר סידורי", anchor="e")
        serial_number_label_window.grid(row=2, column=3, sticky="e", padx=(10, 2), pady=2)


        # --- טבלת רשימת הכלים (שורה 3) ---
        tool_list_frame_window = ttkb.Frame(tool_window, padding="15")
        tool_list_frame_window.grid(row=3, column=0, sticky="nsew", padx=10, pady=(0, 10))
        tool_window.grid_rowconfigure(3, weight=1)

        # הגדרת העמודות שיוצגו בטבלת ניהול הכלים (הסרת 'מיקום')
        tool_list_tree_columns_logical = ("מספר סידורי", "תיאור כלי", "שם הכלי")

        tool_list_tree = ttk.Treeview(tool_list_frame_window, columns=tool_list_tree_columns_logical, show="headings")
        tool_list_tree_style = ttk.Style()
        tool_list_tree_style.configure("Treeview.Heading", font=('Helvetica', 12, 'bold'))
        tool_list_tree_style.configure("Treeview", font=('Helvetica', 12), rowheight=25)

        tool_list_tree.tag_configure("evenrow", background="#f8f9fa")
        tool_list_tree.tag_configure("oddrow", background="#e9ecef")

        # הגדרת כותרות ורוחב לעמודות המוצגות
        tool_list_tree.heading("מספר סידורי", text="מספר סידורי", anchor="e")
        tool_list_tree.column("מספר סידורי", width=80, anchor="e")
        tool_list_tree.heading("תיאור כלי", text="תיאור כלי", anchor="e")
        tool_list_tree.column("תיאור כלי", width=150, anchor="e")
        tool_list_tree.heading("שם הכלי", text="שם כלי", anchor="e")
        tool_list_tree.column("שם הכלי", width=150, anchor="e")

        tool_list_tree_scrollbar = ttkb.Scrollbar(tool_list_frame_window, orient="vertical",
                                                  command=tool_list_tree.yview)
        tool_list_tree.configure(yscrollcommand=tool_list_tree_scrollbar.set)

        tool_list_tree_scrollbar.grid(row=0, column=1, sticky="ns")
        tool_list_tree.grid(row=0, column=0, sticky="nsew")

        tool_list_frame_window.grid_columnconfigure(0, weight=1)
        tool_list_frame_window.grid_rowconfigure(0, weight=1)

        tool_window.tool_name_entry = tool_name_entry_window
        tool_window.tool_description_entry = tool_description_entry_window
        tool_window.serial_number_entry = serial_number_entry_window
        tool_window.add_update_button = add_update_button_window
        tool_window.tool_list_tree = tool_list_tree
        tool_window.edit_button = edit_tool_button
        tool_window.delete_button = delete_tool_button
        tool_window.current_tool_editing = None

        add_update_button_window.config(command=lambda: self.add_tool_window(tool_window))
        edit_tool_button.config(command=lambda: self.load_tool_for_editing(tool_window))
        delete_tool_button.config(command=lambda: self.delete_tool_window(tool_window))
        tool_list_tree.bind("<<TreeviewSelect>>", lambda event: self.on_tool_list_select_window(event, tool_window))

        self.populate_tool_list_table_window(tool_list_tree, self.data_manager.tools_data)
        tool_window.transient(self.root)
        tool_window.protocol("WM_DELETE_WINDOW", lambda: self.on_tool_management_window_close(tool_window))


    def on_tool_management_window_close(self, tool_window):
        """מאפס את רפרנס החלון בעת סגירתו."""
        self.tool_management_window = None
        tool_window.destroy()


    def populate_tool_list_table_window(self, tree, data):
        """ממלא את טבלת רשימת הכלים בחלון הניהול בנתונים."""
        for item in tree.get_children():
            tree.delete(item)
        sorted_data = sorted(data, key=lambda tool: tool.get("שם הכלי", ""))
        for i, tool in enumerate(sorted_data):
            tag = 'evenrow' if i % 2 == 0 else 'oddrow'
            # מילוי הטבלה ללא שדה המיקום
            tree.insert("", tk.END, values=(
                tool.get("מספר סידורי", ""), # אינדקס 0
                tool.get("תיאור כלי", ""),   # אינדקס 1
                tool.get("שם הכלי", "")      # אינדקס 2
            ), tags=(tag,))

    def on_tool_list_select_window(self, event, tool_window):
        """מאפשר/מבטל כפתורים ומאפס מצב עריכה בבחירת פריט בניהול כלים."""
        selected_item = tool_window.tool_list_tree.selection()
        if selected_item:
            tool_window.edit_button.config(state=tk.NORMAL)
            tool_window.delete_button.config(state=tk.NORMAL)
            tool_window.tool_name_entry.delete(0, tk.END)
            tool_window.tool_description_entry.delete(0, tk.END)
            tool_window.serial_number_entry.delete(0, tk.END)
            # שדה מיקום הוסר

            tool_window.add_update_button.config(text="הוסף", bootstyle="success", command=lambda: self.add_tool_window(tool_window))
            tool_window.current_tool_editing = None
        else:
            tool_window.edit_button.config(state=tk.DISABLED)
            tool_window.delete_button.config(state=tk.DISABLED)
            tool_window.tool_name_entry.delete(0, tk.END)
            tool_window.tool_description_entry.delete(0, tk.END)
            tool_window.serial_number_entry.delete(0, tk.END)
            # שדה מיקום הוסר

            tool_window.add_update_button.config(text="הוסף", bootstyle="success", command=lambda: self.add_tool_window(tool_window))
            tool_window.current_tool_editing = None

    def load_tool_for_editing(self, tool_window):
        """טוען פרטי כלי נבחר לשדות הקלט לעריכה."""
        selected_item = tool_window.tool_list_tree.selection()
        if not selected_item:
            self.show_toast("אנא בחר כלי מהטבלה לעריכה.", title="בחירה חסרה", bootstyle="warning")
            return
        item_values = tool_window.tool_list_tree.item(selected_item[0], 'values')
        # בדיקת אורך הרשימה לפי מספר העמודות החדש (3)
        if not item_values or len(item_values) < 3:
            self.show_toast("לא ניתן היה לטעון את פרטי הכלי לעריכה.", title="שגיאה", bootstyle="danger")
            return
        # עדכון האינדקסים בהתאם לעמודות החדשות: ("מספר סידורי", "תיאור כלי", "שם הכלי")
        serial_num_from_tree = self.tree_item_value_safe(item_values, 0) # מספר סידורי הוא אינדקס 0
        desc_from_tree = self.tree_item_value_safe(item_values, 1) # תיאור כלי הוא אינדקס 1
        name_from_tree = self.tree_item_value_safe(item_values, 2) # שם הכלי הוא אינדקס 2


        # מציאת אובייקט הכלי המלא בנתונים הפנימיים (ללא תלות במיקום)
        tool_to_edit = self.data_manager.find_tool_by_serial(serial_num_from_tree)
        # אם לא נמצא לפי מספר סידורי, נסה לפי שם ותיאור (אם קיימים מהטבלה)
        if not tool_to_edit and name_from_tree and desc_from_tree is not None:
             tool_to_edit = self.data_manager.find_tool(name_from_tree, desc_from_tree)


        if tool_to_edit:
            tool_window.current_tool_editing = tool_to_edit
            tool_window.tool_name_entry.delete(0, tk.END)
            tool_window.tool_name_entry.insert(0, tool_to_edit.get("שם הכלי", ""))
            tool_window.tool_description_entry.delete(0, tk.END)
            tool_window.tool_description_entry.insert(0, tool_to_edit.get("תיאור כלי", ""))
            tool_window.serial_number_entry.delete(0, tk.END)
            tool_window.serial_number_entry.insert(0, tool_to_edit.get("מספר סידורי", ""))
            # שדה מיקום הוסר

            tool_window.add_update_button.config(text="עדכן", bootstyle="warning", command=lambda: self.update_tool_window(tool_window))
        else:
            self.show_toast("הכלי שנבחר לעריכה לא נמצא בנתונים הפנימיים.", title="שגיאה", bootstyle="danger")
            tool_window.tool_list_tree.selection_set([])
            self.on_tool_list_select_window(None, tool_window)


    def add_tool_window(self, tool_window):
        """מוסיף כלי חדש מחלון הניהול."""
        tool_name = tool_window.tool_name_entry.get().strip()
        tool_description = tool_window.tool_description_entry.get().strip()
        serial_number = tool_window.serial_number_entry.get().strip()
        # location הוסר

        if not tool_name:
            self.show_toast("אנא הזן שם לכלי.", title="קלט חסר", bootstyle="warning")
            return
        if serial_number:
            if self.data_manager.find_tool_by_serial(serial_number):
                 self.show_toast(f"כלי עם המספר הסידורי '{serial_number}' כבר קיים ברשימה.", title="מספר סידורי קיים", bootstyle="warning")
                 return
        else:
            # בדיקה רק אם שם ותיאור שניהם לא ריקים
            if tool_name and tool_description is not None:
                 if self.data_manager.find_tool(tool_name, tool_description):
                      self.show_toast(f"כלי עם השם '{tool_name}' והתיאור '{tool_description}' כבר קיים ברשימה (ולא קיים לו מספר סידורי ייחודי).", title="כלי קיים", bootstyle="warning")
                      return


        new_tool = {
            "שם הכלי": tool_name, "תיאור כלי": tool_description, "סטטוס": "זמין",
            "שם השואל": "", "תאריך השאלה": "", "מונה השאלות": "0",
            "מספר סידורי": serial_number # 'מיקום' הוסר
        }
        self.data_manager.tools_data.append(new_tool)
        self.data_manager.save_all_data()
        self.populate_tool_list_table_window(tool_window.tool_list_tree, self.data_manager.tools_data)
        self.filter_tools(None)
        tool_window.tool_name_entry.delete(0, tk.END)
        tool_window.tool_description_entry.delete(0, tk.END)
        tool_window.serial_number_entry.delete(0, tk.END)
        # שדה מיקום הוסר

        self.show_toast(f"הכלי '{tool_name}' נוסף בהצלחה.", title="הוספת כלי", bootstyle="success")


    def update_tool_window(self, tool_window):
        """מעדכן פרטי כלי נערך בחלון הניהול."""
        tool_to_update = tool_window.current_tool_editing
        new_name = tool_window.tool_name_entry.get().strip()
        new_description = tool_window.tool_description_entry.get().strip()
        new_serial_number = tool_window.serial_number_entry.get().strip()
        # new_location הוסר


        if not tool_to_update:
            self.show_toast("אין כלי שנבחר לעדכון.", title="שגיאה", bootstyle="warning")
            self.on_tool_list_select_window(None, tool_window)
            return
        if not new_name:
            self.show_toast("אנא הזן שם לכלי.", title="קלט חסר", bootstyle="warning")
            return

        # בדיקה אם המספר הסידורי החדש קיים כבר אצל כלי אחר
        if new_serial_number and new_serial_number != tool_to_update.get("מספר סידורי", ""):
             for tool in self.data_manager.tools_data:
                  if tool is not tool_to_update and tool.get("מספר סידורי", "") == new_serial_number:
                       self.show_toast(f"כלי אחר עם המספר הסידורי '{new_serial_number}' כבר קיים ברשימה.", title="מספר סידורי קיים", bootstyle="warning")
                       return

        # בדיקה אם שם + תיאור קיימים כבר אצל כלי אחר (רק אם אין מספר סידורי חדש או שהמספר הסידורי לא השתנה)
        if not new_serial_number or (new_serial_number == tool_to_update.get("מספר סידורי", "")):
            # בדיקה רק אם שם ותיאור שניהם לא ריקים
            if new_name and new_description is not None:
                for tool in self.data_manager.tools_data:
                    if tool is not tool_to_update and tool.get('שם הכלי') == new_name and tool.get('תיאור כלי', '') == new_description:
                        # אם הכלי הקיים האחר ללא מספר סידורי - זו התנגשות
                        if not tool.get("מספר סידורי"):
                            self.show_toast(f"כלי אחר עם השם '{new_name}' והתיאור '{new_description}' כבר קיים ברשימה (ולא קיים לו מספר סידורי ייחודי).", title="כלי קיים", bootstyle="warning")
                            return


        tool_to_update["שם הכלי"] = new_name
        tool_to_update["תיאור כלי"] = new_description
        tool_to_update["מספר סידורי"] = new_serial_number
        # שדה מיקום הוסר


        self.data_manager.save_all_data()
        self.populate_tool_list_table_window(tool_window.tool_list_tree, self.data_manager.tools_data)
        self.filter_tools(None)
        self.show_toast("הכלי עודכן בהצלחה.", title="עדכון כלי", bootstyle="success")
        self.on_tool_list_select_window(None, tool_window)


    def delete_tool_window(self, tool_window):
        """מוחק כלי נבחר מחלון הניהול."""
        selected_item = tool_window.tool_list_tree.selection()
        if not selected_item:
            self.show_toast("אנא בחר כלי מהטבלה למחיקה.", title="בחירה חסרה", bootstyle="warning")
            return
        item_values = tool_window.tool_list_tree.item(selected_item[0], 'values')
        # בדיקת אורך הרשימה לפי מספר העמודות החדש (3)
        if not item_values or len(item_values) < 3:
            self.show_toast("לא ניתן היה לזהות את פרטי הכלי למחיקה.", title="שגיאה", bootstyle="danger")
            return

        # עדכון האינדקסים בהתאם לעמודות החדשות: ("מספר סידורי", "תיאור כלי", "שם הכלי")
        serial_num_from_tree = self.tree_item_value_safe(item_values, 0) # אינדקס 0
        desc_from_tree = self.tree_item_value_safe(item_values, 1) # אינדקס 1
        name_from_tree = self.tree_item_value_safe(item_values, 2) # אינדקס 2


        tool_to_delete = self.data_manager.find_tool_by_serial(serial_num_from_tree)
        # אם לא נמצא לפי מספר סידורי, נסה לפי שם ותיאור (אם קיימים מהטבלה)
        if not tool_to_delete and name_from_tree and desc_from_tree is not None:
             tool_to_delete = self.data_manager.find_tool(name_from_tree, desc_from_tree)


        if not tool_to_delete:
            self.show_toast("הכלי שנבחר למחיקה לא נמצא בנתונים הפנימיים.", title="שגיאה", bootstyle="danger")
            self.on_tool_list_select_window(None, tool_window)
            return
        if tool_to_delete.get("סטטוס") == "מושאל":
            borrower_name = tool_to_delete.get("שם השואל", "")
            tool_id_str = f"{tool_to_delete.get('שם הכלי', '')} ({tool_to_delete.get('תיאור כלי', '')})"
            if tool_to_delete.get('מספר סידורי'): tool_id_str += f" [{tool_to_delete.get('מספר סידורי')}]"
            self.show_toast(f"לא ניתן למחוק את הכלי '{tool_id_str}' מכיוון שהוא מושאל לשואל '{borrower_name}'.", title="שגיאה", bootstyle="warning")
            return
        confirm = messagebox.askyesno("אישור מחיקה", f"האם אתה בטוח שברצונך למחוק את הכלי '{tool_to_delete.get('שם הכלי', '')} ({tool_to_delete.get('תיאור כלי', '')})'? פעולה זו אינה הפיכה (למעט באמצעות ביטול כל הפעולות).")
        if confirm:
            try:
                # הסרת הכלי מרשימת הנתונים
                self.data_manager.tools_data.remove(tool_to_delete)
            except ValueError:
                 self.show_toast("אירעה שגיאה בעת מחיקת הכלי - הכלי לא נמצא ברשימה הפנימית.", title="שגיאה פנימית", bootstyle="danger")
                 return
            except Exception as e:
                self.show_toast(f"אירעה שגיאה בעת מחיקת הכלי: {e}", title="שגיאה פנימית", bootstyle="danger")
                return

            self.data_manager.save_all_data()
            self.populate_tool_list_table_window(tool_window.tool_list_tree, self.data_manager.tools_data)
            self.filter_tools(None)
            tool_id_str = f"{tool_to_delete.get('שם הכלי', '')} ({tool_to_delete.get('תיאור כלי', '')})"
            if tool_to_delete.get('מספר סידורי'): tool_id_str += f" [{tool_to_delete.get('מספר סידורי')}]"
            self.show_toast(f"הכלי '{tool_id_str}' נמחק בהצלחה.", title="מחיקת כלי", bootstyle="success")
            # אם הכלי הנמחק היה במצב עריכה, אאפס את מצב העריכה בחלון הניהול
            if tool_window.current_tool_editing is tool_to_delete:
                 self.on_tool_list_select_window(None, tool_window)


    def show_manage_borrowers_window(self):
        """פותח חלון חדש לניהול שואלים."""
         # בדיקה אם החלון כבר פתוח
        if hasattr(self, 'borrower_management_window') and self.borrower_management_window is not None and self.borrower_management_window.winfo_exists():
            self.borrower_management_window.lift() # הבאת החלון לפורגראונד
            return

        borrower_window = ttkb.Toplevel(self.root)
        borrower_window.title("ניהול שואלים")

        # --- הוספת קוד למיקום החלון במרכז המסך ---
        window_width = 880  # רוחב החלון (ניתן להתאים) - השארנו את הרוחב המקורי שהיה מוגדר
        window_height = 600 # גובה החלון (ניתן להתאים)
        screen_width = borrower_window.winfo_screenwidth()
        screen_height = borrower_window.winfo_screenheight()
        center_x = int(screen_width / 2 - window_width / 2)
        center_y = int(screen_height / 2 - window_height / 2)
        borrower_window.geometry(f"{window_width}x{window_height}+{center_x}+{center_y}")
        # --------------------------------------------------

        # שמירת רפרנס לחלון הניהול
        self.borrower_management_window = borrower_window

        borrower_controls_frame = ttkb.Frame(borrower_window, padding="15")
        borrower_controls_frame.grid(row=0, column=0, sticky="ew", padx=10, pady=5)
        borrower_window.grid_columnconfigure(0, weight=1)

        # --- פריסת grid ישירה בתוך borrower_controls_frame (כמו בחלון הכלים) ---
        # הגדרת עמודות: כפתור (0), שדה קלט (1), רווח (2), תווית (3)
        borrower_controls_frame.grid_columnconfigure(0, weight=0) # כפתור - לא תתרחב
        borrower_controls_frame.grid_columnconfigure(1, weight=1) # שדה קלט - יתרחב
        borrower_controls_frame.grid_columnconfigure(2, weight=0) # רווח קטן - לא יתרחב (אפשר גם להוריד אם לא נחוץ)
        borrower_controls_frame.grid_columnconfigure(3, weight=0) # תווית - לא תתרחב


        # שורה 0: הוסף שואל + שם שואל
        add_update_button_window = ttkb.Button(borrower_controls_frame, text="הוסף", bootstyle="success")
        # שינוי מיקום: עמודה 0, שורה 0
        add_update_button_window.grid(row=0, column=0, sticky="w", padx=5, pady=2)

        borrower_name_entry_window = ttkb.Entry(borrower_controls_frame, width=20, justify='right')
         # שינוי מיקום: עמודה 1, שורה 0
        borrower_name_entry_window.grid(row=0, column=1, sticky="ew", padx=5, pady=2)
        borrower_name_entry_window.bind("<Escape>", self.clear_entry_on_escape)


        borrower_name_label_window = ttkb.Label(borrower_controls_frame, text=":שם", anchor="e")
        # שינוי מיקום: עמודה 3, שורה 0
        borrower_name_label_window.grid(row=0, column=3, sticky="e", padx=(10, 2), pady=2)


        # שורה 1: ערוך שואל + טלפון שואל
        edit_borrower_button = ttkb.Button(borrower_controls_frame, text="ערוך", bootstyle="info", state=tk.DISABLED)
        # שינוי מיקום: עמודה 0, שורה 1
        edit_borrower_button.grid(row=1, column=0, sticky="w", padx=5, pady=2)

        borrower_phone_entry_window = ttkb.Entry(borrower_controls_frame, width=14, justify='right')

        # שינוי מיקום: עמודה 1, שורה 1
        borrower_phone_entry_window.grid(row=1, column=1, sticky="ew", padx=5, pady=2)
        borrower_phone_entry_window.bind("<Escape>", self.clear_entry_on_escape)


        borrower_phone_label_window = ttkb.Label(borrower_controls_frame, text=":טלפון", anchor="e")
        # שינוי מיקום: עמודה 3, שורה 1
        borrower_phone_label_window.grid(row=1, column=3, sticky="e", padx=(10, 2), pady=2)


        # שורה 2: מחק שואל + כתובת שואל
        delete_borrower_button = ttkb.Button(borrower_controls_frame, text="מחק", bootstyle="danger", state=tk.DISABLED)
        # שינוי מיקום: עמודה 0, שורה 2
        delete_borrower_button.grid(row=2, column=0, sticky="w", padx=5, pady=2)

        borrower_address_entry_window = ttkb.Entry(borrower_controls_frame, width=20, justify='right')
        # שינוי מיקום: עמודה 1, שורה 2
        borrower_address_entry_window.grid(row=2, column=1, sticky="ew", padx=5, pady=2)
        borrower_address_entry_window.bind("<Escape>", self.clear_entry_on_escape)


        borrower_address_label_window = ttkb.Label(borrower_controls_frame, text=":כתובת", anchor="e")
        # שינוי מיקום: עמודה 3, שורה 2
        borrower_address_label_window.grid(row=2, column=3, sticky="e", padx=(10, 2), pady=2)


        # --- טבלת רשימת השואלים (שורה 3) - המיקום של הטבלה נשאר זהה ---
        borrower_list_frame_window = ttkb.Frame(borrower_window, padding="15")
        borrower_list_frame_window.grid(row=1, column=0, sticky="nsew", padx=10, pady=(0, 10))
        borrower_window.grid_rowconfigure(1, weight=1)

        borrower_list_tree_columns_logical = ("כתובת", "מספר טלפון", "שם השואל")

        borrower_list_tree = ttk.Treeview(borrower_list_frame_window, columns=borrower_list_tree_columns_logical, show="headings")
        borrower_list_tree_style = ttk.Style()
        borrower_list_tree_style.configure("Treeview.Heading", font=('Helvetica', 12, 'bold'))
        borrower_list_tree_style.configure("Treeview", font=('Helvetica', 12), rowheight=25)

        borrower_list_tree.tag_configure("evenrow", background="#f8f9fa")
        borrower_list_tree.tag_configure("oddrow", background="#e9ecef")

        borrower_list_tree.heading("שם השואל", text="שם שואל", anchor="e")
        borrower_list_tree.column("שם השואל", width=150, anchor="e")
        borrower_list_tree.heading("מספר טלפון", text="מספר טלפון", anchor="e")
        borrower_list_tree.column("מספר טלפון", width=150, anchor="e")
        borrower_list_tree.heading("כתובת", text="כתובת", anchor="e")
        borrower_list_tree.column("כתובת", width=250, anchor="e")

        borrower_list_tree_scrollbar = ttkb.Scrollbar(borrower_list_frame_window, orient="vertical", command=borrower_list_tree.yview)
        borrower_list_tree.configure(yscrollcommand=borrower_list_tree_scrollbar.set)

        borrower_list_tree_scrollbar.grid(row=0, column=1, sticky="ns")
        borrower_list_tree.grid(row=0, column=0, sticky="nsew")

        borrower_list_frame_window.grid_columnconfigure(0, weight=1)
        borrower_list_frame_window.grid_rowconfigure(0, weight=1)

        # Assign widgets to borrower_window object for easier access in other methods
        borrower_window.borrower_name_entry = borrower_name_entry_window
        borrower_window.borrower_phone_entry = borrower_phone_entry_window
        borrower_window.borrower_address_entry = borrower_address_entry_window
        borrower_window.add_update_button = add_update_button_window
        borrower_window.borrower_list_tree = borrower_list_tree
        borrower_window.edit_button = edit_borrower_button
        borrower_window.delete_button = delete_borrower_button
        borrower_window.current_borrower_editing = None

        # Configure button commands
        add_update_button_window.config(command=lambda: self.add_borrower_window(borrower_window))
        edit_borrower_button.config(command=lambda: self.load_borrower_for_editing(borrower_window))
        delete_borrower_button.config(command=lambda: self.delete_borrower_window(borrower_window))

        # Bind selection event
        borrower_list_tree.bind("<<TreeviewSelect>>", lambda event: self.on_borrower_list_select_window(event, borrower_window))

        self.populate_borrower_list_table_window(borrower_list_tree, self.data_manager.borrowers_data)

        # Set window transient and protocol for closing
        borrower_window.transient(self.root)
        borrower_window.protocol("WM_DELETE_WINDOW", lambda: self.on_borrower_management_window_close(borrower_window))

    def on_borrower_management_window_close(self, borrower_window):
        """מאפס את רפרנס החלון בעת סגירתו."""
        self.borrower_management_window = None
        borrower_window.destroy()


    def populate_borrower_list_table_window(self, tree, data):
        """ממלא את טבלת רשימת השואלים בחלון הניהול."""
        for item in tree.get_children():
            tree.delete(item)
        sorted_data = sorted(data, key=lambda borrower: borrower.get("שם השואל", ""))
        for i, borrower in enumerate(sorted_data):
            tag = 'evenrow' if i % 2 == 0 else 'oddrow'
            tree.insert("", tk.END, values=(
                borrower.get("כתובת", ""),
                borrower.get("מספר טלפון", ""),
                borrower.get("שם השואל", "")
            ), tags=(tag,))

    def on_borrower_list_select_window(self, event, borrower_window):
        """מאפשר/מבטל כפתורים ומאפס מצב עריכה בבחירת פריט בניהול שואלים."""
        selected_item = borrower_window.borrower_list_tree.selection()
        if selected_item:
            borrower_window.edit_button.config(state=tk.NORMAL)
            borrower_window.delete_button.config(state=tk.NORMAL)
            borrower_window.borrower_name_entry.delete(0, tk.END)
            borrower_window.borrower_phone_entry.delete(0, tk.END)
            borrower_window.borrower_address_entry.delete(0, tk.END)
            borrower_window.add_update_button.config(text="הוסף", bootstyle="success", command=lambda: self.add_borrower_window(borrower_window))
            borrower_window.current_borrower_editing = None
        else:
            borrower_window.edit_button.config(state=tk.DISABLED)
            borrower_window.delete_button.config(state=tk.DISABLED)
            borrower_window.borrower_name_entry.delete(0, tk.END)
            borrower_window.borrower_phone_entry.delete(0, tk.END)
            borrower_window.borrower_address_entry.delete(0, tk.END)
            borrower_window.add_update_button.config(text="הוסף", bootstyle="success", command=lambda: self.add_borrower_window(borrower_window))
            borrower_window.current_borrower_editing = None

    def load_borrower_for_editing(self, borrower_window):
        """טוען פרטי שואל נבחר לשדות הקלט לעריכה."""
        selected_item = borrower_window.borrower_list_tree.selection()
        if not selected_item:
            self.show_toast("אנא בחר שואל מהטבלה לעריכה.", title="בחירה חסרה", bootstyle="warning")
            return
        item_values = borrower_window.borrower_list_tree.item(selected_item[0], 'values')
        # אינדקס 2 הוא שם השואל בטבלת ניהול השואלים
        if not item_values or len(item_values) < 3:
            self.show_toast("לא ניתן היה לטעון את פרטי השואל לעריכה.", title="שגיאה", bootstyle="danger")
            return
        borrower_name = self.tree_item_value_safe(item_values, 2)
        borrower_to_edit = self.data_manager.find_borrower(borrower_name)
        if borrower_to_edit:
            borrower_window.current_borrower_editing = borrower_to_edit
            borrower_window.borrower_name_entry.delete(0, tk.END)
            borrower_window.borrower_name_entry.insert(0, borrower_to_edit.get("שם השואל", ""))
            borrower_window.borrower_phone_entry.delete(0, tk.END)
            borrower_window.borrower_phone_entry.insert(0, borrower_to_edit.get("מספר טלפון", ""))
            borrower_window.borrower_address_entry.delete(0, tk.END)
            borrower_window.borrower_address_entry.insert(0, borrower_to_edit.get("כתובת", ""))
            borrower_window.add_update_button.config(text="עדכן", bootstyle="warning", command=lambda: self.update_borrower_window(borrower_window))
        else:
            self.show_toast("השואל שנבחר לעריכה לא נמצא בנתונים הפנימיים.", title="שגיאה", bootstyle="danger")
            borrower_window.borrower_list_tree.selection_set([])
            self.on_borrower_list_select_window(None, borrower_window)


    def add_borrower_window(self, borrower_window):
        """מוסיף שואל חדש מחלון הניהול."""
        borrower_name = borrower_window.borrower_name_entry.get().strip()
        borrower_phone = borrower_window.borrower_phone_entry.get().strip()
        borrower_address = borrower_window.borrower_address_entry.get().strip()
        if not borrower_name:
            self.show_toast("אנא הזן שם לשואל.", title="קלט חסר", bootstyle="warning")
            return
        if self.data_manager.find_borrower(borrower_name):
            self.show_toast(f"שואל עם השם '{borrower_name}' כבר קיים ברשימה.", title="שואל קיים", bootstyle="warning")
            return
        new_borrower = {
            "שם השואל": borrower_name, "מספר טלפון": borrower_phone, "כתובת": borrower_address
        }
        self.data_manager.borrowers_data.append(new_borrower)
        self.data_manager.save_all_data()
        self.populate_borrower_list_table_window(borrower_window.borrower_list_tree, self.data_manager.borrowers_data)
        self.filter_borrower_table(None)
        borrower_window.borrower_name_entry.delete(0, tk.END)
        borrower_window.borrower_phone_entry.delete(0, tk.END)
        borrower_window.borrower_address_entry.delete(0, tk.END)
        self.show_toast(f"השואל '{borrower_name}' נוסף בהצלחה.", title="הוספת שואל", bootstyle="success")

    def update_borrower_window(self, borrower_window):
        """מעדכן פרטי שואל נערך בחלון הניהול."""
        borrower_to_update = borrower_window.current_borrower_editing
        new_name = borrower_window.borrower_name_entry.get().strip()
        new_phone = borrower_window.borrower_phone_entry.get().strip()
        new_address = borrower_window.borrower_address_entry.get().strip()

        if not borrower_to_update:
            self.show_toast("אין שואל שנבחר לעדכון.", title="שגיאה", bootstyle="warning")
            self.on_borrower_list_select_window(None, borrower_window)
            return
        if not new_name:
            self.show_toast("אנא הזן שם לשואל.", title="קלט חסר", bootstyle="warning")
            return

        original_name = borrower_to_update.get("שם השואל")

        for borrower in self.data_manager.borrowers_data:
            if borrower is not borrower_to_update and borrower.get('שם השואל') == new_name:
                self.show_toast(f"שואל עם השם '{new_name}' כבר קיים ברשימה.", title="שואל קיים", bootstyle="warning")
                return

        borrower_to_update["שם השואל"] = new_name
        borrower_to_update["מספר טלפון"] = new_phone
        borrower_to_update["כתובת"] = new_address

        # עדכון שם השואל גם בכלי מושאלים אם הוא היה משויך לשואל
        if original_name and original_name != new_name:
             for tool in self.data_manager.tools_data:
                  if tool.get("שם השואל") == original_name:
                       tool["שם השואל"] = new_name

        self.data_manager.save_all_data()
        self.populate_borrower_list_table_window(borrower_window.borrower_list_tree, self.data_manager.borrowers_data)
        self.refresh_ui() # רענון ממשק ראשי כדי לעדכן את שם השואל בטבלת הכלים הראשית
        self.show_toast("השואל עודכן בהצלחה.", title="עדכון שואל", bootstyle="success")
        self.on_borrower_list_select_window(None, borrower_window)

    def delete_borrower_window(self, borrower_window):
        """מוחק שואל נבחר מחלון הניהול."""
        selected_item = borrower_window.borrower_list_tree.selection()
        if not selected_item:
            self.show_toast("אנא בחר שואל מהטבלה למחיקה.", title="בחירה חסרה", bootstyle="warning")
            return
        item_values = borrower_window.borrower_list_tree.item(selected_item[0], 'values')
        # אינדקס 2 הוא שם השואל בטבלת ניהול השואלים
        if not item_values or len(item_values) < 3:
            self.show_toast("לא ניתן היה לזהות את פרטי השואל למחיקה.", title="שגיאה", bootstyle="danger")
            return
        borrower_name = self.tree_item_value_safe(item_values, 2)
        borrower_to_delete = self.data_manager.find_borrower(borrower_name)
        if not borrower_to_delete:
            self.show_toast("השואל שנבחר למחיקה לא נמצא בנתונים הפנימיים.", title="שגיאה", bootstyle="danger")
            self.on_borrower_list_select_window(None, borrower_window)
            return
        borrowed_tools_by_this_borrower = [tool for tool in self.data_manager.tools_data if tool.get("שם השואל") == borrower_name and tool.get("סטטוס") == "מושאל"]
        if borrowed_tools_by_this_borrower:
            self.show_toast(f"לא ניתן למחוק את השואל '{borrower_name}' מכיוון שיש לו כלים מושאלים. אנא החזר את הכלים לפני המחיקה.", title="שגיאה", bootstyle="warning")
            return
        confirm = messagebox.askyesno("אישור מחיקה", f"האם אתה בטוח שברצונך למחוק את השואל '{borrower_name}'? פעולה זו אינה הפיכה (למעט באמצעות ביטול כל הפעולות).")
        if confirm:
            try:
                # הסרת השואל מרשימת הנתונים
                self.data_manager.borrowers_data.remove(borrower_to_delete)
            except ValueError:
                 self.show_toast("אירעה שגיאה בעת מחיקת השואל - השואל לא נמצא ברשימה הפנימית.", title="שגיאה פנימית", bootstyle="danger")
                 return
            except Exception as e:
                self.show_toast(f"אירעה שגיאה בעת מחיקת השואל: {e}", title="שגיאה פנימית", bootstyle="danger")
                return

            self.data_manager.save_all_data()
            self.populate_borrower_list_table_window(borrower_window.borrower_list_tree, self.data_manager.borrowers_data)
            self.refresh_ui() # רענון ממשק ראשי כדי להסיר את שם השואל אם היה נבחר
            self.show_toast(f"השואל '{borrower_name}' נמחק בהצלחה.", title="מחיקת שואל", bootstyle="success")
            # אם השואל הנמחק היה במצב עריכה, אאפס את מצב העריכה בחלון הניהול
            if borrower_window.current_borrower_editing is borrower_to_delete:
                 self.on_borrower_list_select_window(None, borrower_window)


# --- הפעלת היישום ---
if __name__ == "__main__":
    root = ttkb.Window(themename="flatly")
    app = ToolBorrowingApp(root)
    root.place_window_center()
    root.mainloop()