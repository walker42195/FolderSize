# Skriven av Fredrik Sandgren och AI 2025-08-07
import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading
import pandas as pd
import matplotlib.pyplot as plt
import subprocess

WARNING_SIZE = 1 * 1024**3  # 1 GB

def get_folder_info(path):
    total_size = 0
    file_count = 0
    try:
        for root, dirs, files in os.walk(path, topdown=True):
            for f in files:
                try:
                    fp = os.path.join(root, f)
                    if os.path.exists(fp):  # Extra kontroll
                        total_size += os.path.getsize(fp)
                        file_count += 1
                except (OSError, UnicodeError, PermissionError) as e:
                    print(f"Kunde inte läsa fil {f}: {e}")
                    continue
    except (OSError, UnicodeError, PermissionError) as e:
        print(f"Kunde inte läsa mapp {path}: {e}")
    return total_size, file_count

def sizeof_fmt(num, suffix="B"):
    for unit in ["", "K", "M", "G", "T"]:
        if abs(num) < 1024.0:
            return f"{num:3.1f} {unit}{suffix}"
        num /= 1024.0
    return f"{num:.1f} P{suffix}"

class FolderViewer(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Mappstorleksvisare")
        self.geometry("1000x700")

        self.folder_data = []
        self.total_bytes = 0
        self.root_path = ""
        
        # Sorteringsvariabler
        self.sort_column = None
        self.sort_reverse = False

        # Sök och knappar
        top_frame = tk.Frame(self)
        top_frame.pack(fill="x")

        tk.Label(top_frame, text="Sök:").pack(side="left", padx=5)
        self.search_var = tk.StringVar()
        self.search_var.trace_add("write", self.filter_tree)
        search_entry = tk.Entry(top_frame, textvariable=self.search_var)
        search_entry.pack(side="left", padx=5)
        
        tk.Button(top_frame, text="Välj mapp", command=self.select_folder).pack(side="left", padx=5)
        tk.Button(top_frame, text="Exportera CSV", command=self.export_csv).pack(side="right", padx=5)
        tk.Button(top_frame, text="Exportera Excel", command=self.export_excel).pack(side="right")
        tk.Button(top_frame, text="Visa diagram", command=self.show_pie_chart).pack(side="right")

        # Trädvy med sortering
        self.tree = ttk.Treeview(self, columns=("size", "files"), show="tree headings")
        self.tree.heading("#0", text="Mapp / Fil", command=lambda: self.sort_treeview("#0"))
        self.tree.heading("size", text="Storlek ↕", command=lambda: self.sort_treeview("size"))
        self.tree.heading("files", text="Filer ↕", command=lambda: self.sort_treeview("files"))
        self.tree.column("size", width=100, anchor="e")
        self.tree.column("files", width=70, anchor="center")
        self.tree.pack(fill="both", expand=True)

        self.tree.bind("<<TreeviewOpen>>", self.on_open)
        self.tree.bind("<Button-3>", self.show_context_menu)  # Högerklick
        self.context_menu = tk.Menu(self, tearoff=0)
        self.context_menu.add_command(label="Öppna i Explorer", command=self.open_in_explorer)

        self.progress = ttk.Progressbar(self, mode="indeterminate")
        self.progress.pack(fill="x")

        self.status = tk.Label(self, text="Totalt: -", anchor="w")
        self.status.pack(fill="x")

        self.after(100, self.select_folder)

    def sort_treeview(self, column):
        """Sortera trädvyn baserat på vald kolumn"""
        # Om samma kolumn klickas igen, växla riktning
        if self.sort_column == column:
            self.sort_reverse = not self.sort_reverse
        else:
            self.sort_column = column
            self.sort_reverse = False
        
        # Uppdatera kolumnrubriker med pilar
        self.update_column_headers()
        
        # Sortera alla nivåer rekursivt
        self.sort_children("", column)
    
    def sort_children(self, parent, column):
        """Sortera barn till en viss nod rekursivt"""
        children = list(self.tree.get_children(parent))
        if not children:
            return
        
        # Sortera baserat på kolumn
        if column == "#0":
            # Sortera på mappnamn (alfabetiskt)
            children.sort(key=lambda x: self.tree.item(x, "text").lower(), 
                         reverse=self.sort_reverse)
        elif column == "size":
            # Sortera på storlek
            def get_size_value(item_id):
                values = self.tree.item(item_id, "values")
                if not values or len(values) == 0:
                    return 0
                size_str = values[0]
                if size_str and size_str != "Beräknar..." and size_str != "":
                    # Konvertera tillbaka från formaterad sträng till bytes för korrekt sortering
                    try:
                        if "G" in size_str:
                            return float(size_str.split()[0]) * 1024**3
                        elif "M" in size_str:
                            return float(size_str.split()[0]) * 1024**2
                        elif "K" in size_str:
                            return float(size_str.split()[0]) * 1024
                        else:
                            return float(size_str.split()[0])
                    except (ValueError, IndexError):
                        return 0
                return 0
            children.sort(key=get_size_value, reverse=self.sort_reverse)
        elif column == "files":
            # Sortera på antal filer
            def get_files_value(item_id):
                values = self.tree.item(item_id, "values")
                if not values or len(values) < 2:
                    return 0
                files_str = values[1]
                if files_str and files_str != "—" and files_str != "Beräknar...":
                    try:
                        return int(files_str)
                    except ValueError:
                        return 0
                return 0
            children.sort(key=get_files_value, reverse=self.sort_reverse)
        
        # Flytta alla objekt i sorterad ordning
        for index, item in enumerate(children):
            self.tree.move(item, parent, index)
        
        # Sortera barn till varje barn (rekursivt)
        for child in children:
            if self.tree.get_children(child):
                self.sort_children(child, column)
    
    def update_column_headers(self):
        """Uppdatera kolumnrubriker med sorterings-pilar"""
        # Återställ alla rubriker
        self.tree.heading("#0", text="Mapp / Fil")
        self.tree.heading("size", text="Storlek ↕")
        self.tree.heading("files", text="Filer ↕")
        
        # Lägg till pil för aktiv sorteringskolumn
        arrow = " ↓" if self.sort_reverse else " ↑"
        
        if self.sort_column == "#0":
            self.tree.heading("#0", text="Mapp / Fil" + arrow)
        elif self.sort_column == "size":
            self.tree.heading("size", text="Storlek" + arrow)
        elif self.sort_column == "files":
            self.tree.heading("files", text="Filer" + arrow)

    def select_folder(self):
        path = filedialog.askdirectory(title="Välj mapp att skanna")
        if not path:
            return
        self.root_path = path
        self.tree.delete(*self.tree.get_children())
        self.folder_data.clear()
        self.total_bytes = 0
        self.status.config(text="Totalt: -")
        
        # Återställ sortering
        self.sort_column = None
        self.sort_reverse = False
        self.update_column_headers()
        
        self.insert_node("", path)

    def insert_node(self, parent, path):
        try:
            node_text = os.path.basename(path) or path
            node = self.tree.insert(parent, "end", text=node_text,
                                    values=("Beräknar...", "—"), open=False,
                                    tags=(path,))  # Här sparas full path som tagg

            try:
                # Kontrollera om mappen är läsbar innan vi lägger till dummy
                subdirs = []
                for item in os.listdir(path):
                    item_path = os.path.join(path, item)
                    if os.path.isdir(item_path):
                        subdirs.append(item)
                
                if subdirs:
                    self.tree.insert(node, "end", text="", values=("", ""))  # Dummy med tomma värden
            except (OSError, PermissionError, UnicodeError) as e:
                print(f"Kunde inte kontrollera undermappar i {path}: {e}")

            self.progress.start()
            threading.Thread(target=self.update_node_info, args=(node, path), daemon=True).start()
            
        except (UnicodeError, OSError) as e:
            print(f"Kunde inte lägga till nod för {path}: {e}")
            return None
        
        return node

    def populate_node(self, node, path):
        try:
            # Försök läsa mappen med olika kodningar
            try:
                entries = os.listdir(path)
            except UnicodeDecodeError:
                # Försök med olika kodningar
                try:
                    entries = os.listdir(path.encode('utf-8').decode('utf-8'))
                except:
                    print(f"Kunde inte läsa mapp: {path}")
                    return
            
            entries = sorted(entries, key=str.lower)  # Case-insensitive sortering
        except (OSError, PermissionError, UnicodeError) as e:
            print(f"Åtkomst nekad till mapp {path}: {e}")
            return

        for entry in entries:
            try:
                full_path = os.path.join(path, entry)
                
                # Kontrollera att filen/mappen verkligen existerar
                if not os.path.exists(full_path):
                    continue
                    
                if os.path.isdir(full_path):
                    self.insert_node(node, full_path)
                else:
                    try:
                        size = os.path.getsize(full_path)
                        size_str = sizeof_fmt(size)
                        file_node = self.tree.insert(node, "end", text=entry, 
                                                   values=(size_str, ""), 
                                                   tags=(full_path,), open=False)
                    except (OSError, PermissionError):
                        # Lägg till filen ändå men utan storlek
                        file_node = self.tree.insert(node, "end", text=entry, 
                                                   values=("Åtkomst nekad", ""), 
                                                   tags=(full_path,), open=False)
            except (UnicodeError, OSError, PermissionError) as e:
                print(f"Problem med {entry}: {e}")
                continue

    def update_node_info(self, node, path):
        size, count = get_folder_info(path)
        self.total_bytes += size

        display_size = sizeof_fmt(size)
        self.tree.set(node, "size", display_size)
        self.tree.set(node, "files", str(count))

        if size >= WARNING_SIZE:
            # Lägg till warning-tagg utan att påverka path-taggen
            current_tags = list(self.tree.item(node, "tags"))
            if "warning" not in current_tags:
                current_tags.append("warning")
                self.tree.item(node, tags=current_tags)
            self.tree.tag_configure("warning", foreground="red")

        full_path = self.get_full_path(node)
        if full_path:  # Kontrollera att vi har en giltig path
            self.folder_data.append({
                "Path": full_path,
                "Size": size,
                "SizeStr": display_size,
                "Files": count
            })

        self.status.config(text=f"Totalt: {sizeof_fmt(self.total_bytes)}")
        self.progress.stop()

    def on_open(self, event):
        node = self.tree.focus()
        path = self.get_full_path(node)

        if self.tree.get_children(node):
            first_child = self.tree.get_children(node)[0]
            if not self.tree.item(first_child, "text"):
                self.tree.delete(first_child)
                self.progress.start()
                threading.Thread(target=self.populate_node, args=(node, path), daemon=True).start()

    def get_full_path(self, node):
        tags = self.tree.item(node, "tags")
        if tags:
            # Hitta den första taggen som ser ut som en sökväg (innehåller : eller börjar med /)
            for tag in tags:
                if (":" in tag and len(tag) > 2) or tag.startswith("/"):
                    return tag
        return None

    def filter_tree(self, *args):
        query = self.search_var.get().lower()
        for item in self.tree.get_children(""):
            self.filter_recursive(item, query)

    def filter_recursive(self, node, query):
        label = self.tree.item(node, "text").lower()
        show = query in label

        for child in self.tree.get_children(node):
            if self.filter_recursive(child, query):
                show = True

        self.tree.item(node, open=show)
        self.tree.detach(node) if not show else self.tree.reattach(node, self.tree.parent(node), "end")
        return show

    def export_csv(self):
        if not self.folder_data:
            messagebox.showwarning("Ingen data", "Ingen data att exportera.")
            return
        path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV", "*.csv")])
        if path:
            df = pd.DataFrame(self.folder_data)
            df.to_csv(path, index=False)
            messagebox.showinfo("Exporterat", f"Data exporterad till:\n{path}")

    def export_excel(self):
        if not self.folder_data:
            messagebox.showwarning("Ingen data", "Ingen data att exportera.")
            return
        path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")])
        if path:
            df = pd.DataFrame(self.folder_data)
            df.to_excel(path, index=False)
            messagebox.showinfo("Exporterat", f"Data exporterad till:\n{path}")

    def show_pie_chart(self):
        if not self.folder_data:
            messagebox.showwarning("Ingen data", "Ingen data att visa.")
            return
        sorted_data = sorted(self.folder_data, key=lambda x: x["Size"], reverse=True)[:10]
        labels = [os.path.basename(d["Path"]) or d["Path"] for d in sorted_data]
        sizes = [d["Size"] for d in sorted_data]

        plt.figure(figsize=(8, 6))
        plt.pie(sizes, labels=labels, autopct="%1.1f%%", startangle=140)
        plt.title("Topp 10 största mappar")
        plt.axis("equal")
        plt.show()

    def show_context_menu(self, event):
        node = self.tree.identify_row(event.y)
        if node:
            self.tree.selection_set(node)
            self.context_menu.post(event.x_root, event.y_root)

    def open_in_explorer(self):
        selected = self.tree.selection()
        if not selected:
            return
        node = selected[0]
        path = self.get_full_path(node)
        if path is None:
            messagebox.showerror("Fel", "Ingen sökväg hittades för vald fil eller mapp.")
            return
        
        # Kontrollera att sökvägen existerar
        if not os.path.exists(path):
            messagebox.showerror("Fel", "Filen eller mappen finns inte längre.")
            return
            
        try:
            # Konvertera path till Windows short name för att undvika Unicode-problem
            import ctypes
            from ctypes import wintypes
            
            kernel32 = ctypes.windll.kernel32
            GetShortPathNameW = kernel32.GetShortPathNameW
            GetShortPathNameW.argtypes = [wintypes.LPCWSTR, wintypes.LPWSTR, wintypes.DWORD]
            GetShortPathNameW.restype = wintypes.DWORD
            
            # Få short path
            buffer = ctypes.create_unicode_buffer(260)
            GetShortPathNameW(path, buffer, 260)
            short_path = buffer.value
            
            if short_path:
                if os.path.isfile(short_path):
                    subprocess.run(['explorer', '/select,', short_path], check=True)
                else:
                    subprocess.run(['explorer', short_path], check=True)
            else:
                # Fallback till vanlig metod
                raise Exception("Kunde inte få short path")
                
        except Exception as e:
            try:
                # Fallback 1: os.startfile
                os.startfile(path)
            except Exception as e2:
                try:
                    # Fallback 2: Vanlig explorer-kommando
                    if os.path.isfile(path):
                        subprocess.run(['explorer', '/select,', path], check=True)
                    else:
                        subprocess.run(['explorer', path], check=True)
                except Exception as e3:
                    messagebox.showerror("Fel", f"Kunde inte öppna i Explorer:\n{e}\n{e2}\n{e3}")

if __name__ == "__main__":
    app = FolderViewer()
    app.mainloop()