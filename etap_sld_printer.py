"""
ETAP Automation
"""

import customtkinter as ctk
import threading
import sys
import time
import json
import xml.etree.ElementTree as ET
import pyautogui
import etap.api
import ctypes
import os
from tkinter import filedialog
from collections import defaultdict

# --- WINDOWS API FOCUS HELPER ---
def focus_etap_window():
    """Porta la finestra ETAP in primo piano forzando la visualizzazione."""
    try:
        user32 = ctypes.windll.user32
        found_windows = []

        def callback(hwnd, extra):
            if not user32.IsWindowVisible(hwnd): return True
            length = user32.GetWindowTextLengthW(hwnd)
            if length == 0: return True
            buff = ctypes.create_unicode_buffer(length + 1)
            user32.GetWindowTextW(hwnd, buff, length + 1)
            title = buff.value
            
            # Cerca specificamente "ETAP 24" (es. "ETAP 24.0.3")
            if "ETAP 24" in title and "Shell" not in title and "Python" not in title: 
                found_windows.append((hwnd, title))
            return True

        CMPFUNC = ctypes.CFUNCTYPE(ctypes.c_bool, ctypes.c_int, ctypes.c_int)
        user32.EnumWindows(CMPFUNC(callback), 0)

        target_hwnd = None
        target_title = ""
        
        # Cerca quella che contiene "ETAP 24.0.3"
        for hwnd, title in found_windows:
            if "ETAP 24.0.3" in title:
                target_hwnd = hwnd
                target_title = title
                break
        
        # Fallback: qualsiasi finestra con "ETAP 24"
        if not target_hwnd and found_windows:
            target_hwnd = found_windows[0][0]
            target_title = found_windows[0][1]

        if not target_hwnd:
            print("❌ Nessuna finestra ETAP trovata")
            return False

        print(f"✓ Trovata: '{target_title}'")

        # Salva finestra corrente
        original_hwnd = user32.GetForegroundWindow()
        
        # 1. Ripristina ETAP se minimizzata
        if user32.IsIconic(target_hwnd):
            user32.ShowWindow(target_hwnd, 9)  # SW_RESTORE
            time.sleep(0.3)
        
        # 2. TRUCCO: Minimizza la finestra corrente per forzare il cambio visivo
        if original_hwnd and original_hwnd != target_hwnd:
            print(f"   Minimizzazione temporanea finestra corrente...")
            user32.ShowWindow(original_hwnd, 6)  # SW_MINIMIZE
            time.sleep(0.2)
        
        # 3. Mostra ETAP
        user32.ShowWindow(target_hwnd, 5)  # SW_SHOW
        time.sleep(0.1)

        # 4. Forza Z-order
        SWP_NOMOVE = 0x0002
        SWP_NOSIZE = 0x0001
        SWP_SHOWWINDOW = 0x0040
        HWND_TOPMOST = -1
        HWND_NOTOPMOST = -2
        
        user32.SetWindowPos(target_hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_SHOWWINDOW)
        time.sleep(0.1)
        user32.SetWindowPos(target_hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_SHOWWINDOW)
        time.sleep(0.1)
        
        # 5. Attiva ETAP
        user32.SwitchToThisWindow(target_hwnd, True)
        time.sleep(0.2)
        
        # 6. SetForegroundWindow
        tid_current = user32.GetWindowThreadProcessId(user32.GetForegroundWindow(), 0)
        tid_target = user32.GetWindowThreadProcessId(target_hwnd, 0)
        
        if tid_current != tid_target:
            user32.AttachThreadInput(tid_current, tid_target, True)
            user32.BringWindowToTop(target_hwnd)
            user32.SetForegroundWindow(target_hwnd)
            user32.AttachThreadInput(tid_current, tid_target, False)
        else:
            user32.SetForegroundWindow(target_hwnd)
        
        time.sleep(0.5)

        # Verifica
        is_foreground = user32.GetForegroundWindow() == target_hwnd
        print(f"   Focus ottenuto: {'✓ SI' if is_foreground else '✗ NO (ma dovrebbe essere visibile)'}")
        
        # NON ripristiniamo la finestra originale - vogliamo che ETAP rimanga visibile
        return True  # Ritorna sempre True se l'abbiamo trovata
        
        return False
    except Exception as e:
        print(f"❌ Errore focus: {e}")
        import traceback
        traceback.print_exc()
        return False

# --- CONFIGURAZIONE ---
ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")
BASE_ADDRESS = "https://localhost:60000"

def focus_app_window():
    """MSA (Make Self Active): Porta QUESTA APP in primo piano usando la logica robusta."""
    try:
        user32 = ctypes.windll.user32
        target_title = "ETAP SLD Batch Printer"
        target_hwnd = None
        
        # 1. Trova la finestra
        def callback(hwnd, extra):
            if not user32.IsWindowVisible(hwnd): return True
            length = user32.GetWindowTextLengthW(hwnd)
            buff = ctypes.create_unicode_buffer(length + 1)
            user32.GetWindowTextW(hwnd, buff, length + 1)
            if buff.value == target_title:
                nonlocal target_hwnd
                target_hwnd = hwnd
                return False 
            return True

        CMPFUNC = ctypes.CFUNCTYPE(ctypes.c_bool, ctypes.c_int, ctypes.c_int)
        user32.EnumWindows(CMPFUNC(callback), 0)

        if not target_hwnd: return

        # 2. Minimizza finestra corrente (probabilmente ETAP o altro)
        original_hwnd = user32.GetForegroundWindow()
        if original_hwnd and original_hwnd != target_hwnd:
            user32.ShowWindow(original_hwnd, 6) # SW_MINIMIZE
            time.sleep(0.2)
        
        # 3. Porta su la nostra App
        if user32.IsIconic(target_hwnd):
            user32.ShowWindow(target_hwnd, 9) # SW_RESTORE
            time.sleep(0.1)
            
        user32.ShowWindow(target_hwnd, 5) # SW_SHOW
        
        # Forza Z-order
        user32.SetWindowPos(target_hwnd, -1, 0, 0, 0, 0, 3) # TOPMOST
        user32.SetWindowPos(target_hwnd, -2, 0, 0, 0, 0, 3) # NOTOPMOST
        
        # Attiva
        user32.SwitchToThisWindow(target_hwnd, True)
        
    except: pass


# Cartella di default per l'output
DEFAULT_OUTPUT_DIR = os.path.join(os.getcwd(), "Output_SLD")
if not os.path.exists(DEFAULT_OUTPUT_DIR):
    os.makedirs(DEFAULT_OUTPUT_DIR, exist_ok=True)

CONFIG = {
    "DELAY_KEY_PRESS": 0.2,
    "DELAY_AFTER_SPACE": 2.0,
    "DELAY_BETWEEN_SEQS": 1.0,
    "DELAY_PRINT_DIALOG": 1.5,
    "DELAY_SAVE_DIALOG": 2.0,
    "OUTPUT_DIR": DEFAULT_OUTPUT_DIR
}

class SettingsDialog(ctk.CTkToplevel):
    """Finestra di configurazione parametri"""
    def __init__(self, parent):
        super().__init__(parent)
        self.title("⚙️ Impostazioni")
        self.geometry("600x600")
        self.attributes("-topmost", True)
        
        # Titolo
        ctk.CTkLabel(self, text="Configurazione Parametri", font=("Arial", 18, "bold")).pack(pady=20)
        
        # --- SEZIONE CARTELLA OUTPUT ---
        out_frame = ctk.CTkFrame(self)
        out_frame.pack(fill="x", padx=20, pady=10)
        ctk.CTkLabel(out_frame, text="Cartella Output PDF:", font=("Arial", 12, "bold")).pack(anchor="w", padx=10, pady=5)
        
        self.lbl_path = ctk.CTkEntry(out_frame)
        self.lbl_path.insert(0, CONFIG["OUTPUT_DIR"])
        self.lbl_path.pack(side="left", fill="x", expand=True, padx=10, pady=5)
        
        ctk.CTkButton(out_frame, text="📂 Sfoglia", width=80, command=self.browse_folder).pack(side="right", padx=10)

        # Frame scrollabile per i parametri
        scroll = ctk.CTkScrollableFrame(self, label_text="Ritardi Temporali (secondi)")
        scroll.pack(fill="both", expand=True, padx=20, pady=10)
        
        self.entries = {}
        
        params = [
            ("DELAY_KEY_PRESS", "Ritardo tra pressioni tasti", CONFIG["DELAY_KEY_PRESS"]),
            ("DELAY_AFTER_SPACE", "Attesa dopo Space (refresh grafico)", CONFIG["DELAY_AFTER_SPACE"]),
            ("DELAY_BETWEEN_SEQS", "Pausa tra sequenze di comandi", CONFIG["DELAY_BETWEEN_SEQS"]),
            ("DELAY_PRINT_DIALOG", "Attesa apertura dialogo Stampa", CONFIG["DELAY_PRINT_DIALOG"]),
            ("DELAY_SAVE_DIALOG", "Attesa apertura dialogo Salva", CONFIG["DELAY_SAVE_DIALOG"]),
        ]
        
        for key, label, default in params:
            frame = ctk.CTkFrame(scroll, fg_color="transparent")
            frame.pack(fill="x", pady=8, padx=5)
            
            lbl = ctk.CTkLabel(frame, text=label, anchor="w", width=300)
            lbl.pack(side="left", padx=5)
            
            entry = ctk.CTkEntry(frame, width=80)
            entry.insert(0, str(default))
            entry.pack(side="right", padx=5)
            
            self.entries[key] = entry
        
        # Separatore
        ctk.CTkLabel(self, text="─" * 50, text_color="gray").pack(pady=10)
        
        # Pulsanti
        btn_frame = ctk.CTkFrame(self, fg_color="transparent")
        btn_frame.pack(pady=30)
        
        ctk.CTkButton(btn_frame, text="💾 SALVA", command=self.save, 
                     fg_color="green", width=160, height=45, font=("Arial", 14, "bold")).pack(side="left", padx=15)
        ctk.CTkButton(btn_frame, text="❌ ANNULLA", command=self.destroy, 
                     fg_color="gray", width=160, height=45, font=("Arial", 14, "bold")).pack(side="left", padx=15)
    
    def browse_folder(self):
        d = filedialog.askdirectory(initialdir=CONFIG["OUTPUT_DIR"])
        if d:
            self.lbl_path.delete(0, "end")
            self.lbl_path.insert(0, d)
            self.attributes("-topmost", True) # Riporta in primo piano dopo il dialog

    def save(self):
        try:
            # Salva Path
            new_path = self.lbl_path.get()
            if os.path.isdir(new_path):
                CONFIG["OUTPUT_DIR"] = new_path
            
            # Salva Parametri
            for key, entry in self.entries.items():
                value = float(entry.get().replace(",", "."))
                CONFIG[key] = value
                
            print("✓ Configurazione salvata:", CONFIG)
            self.destroy()
        except ValueError:
            print("❌ Errore: Inserisci solo numeri validi")

class CollapsibleGroup(ctk.CTkFrame):
    """Gruppo espandibile con Checkbox Master e Entry modificabile per Output"""
    def __init__(self, parent, title, scenarios):
        super().__init__(parent, fg_color="transparent")
        self.pack(fill="x", pady=2, padx=2)
        
        self.scenarios = scenarios
        self.is_expanded = True
        self.child_checkboxes = []
        
        # HEADER
        self.header = ctk.CTkFrame(self, fg_color="#2B2B2B", corner_radius=6)
        self.header.pack(fill="x", ipady=5)
        
        self.btn_toggle = ctk.CTkButton(self.header, text="▼", width=30, fg_color="transparent", 
                                        hover_color="#444", command=self.toggle)
        self.btn_toggle.pack(side="left", padx=(5, 0))
        
        self.var_master = ctk.BooleanVar(value=True)
        self.chk_master = ctk.CTkCheckBox(self.header, text=title, variable=self.var_master, 
                                          command=self.on_master_click, font=("Arial", 13, "bold"))
        self.chk_master.pack(side="left", padx=10)
        
        ctk.CTkLabel(self.header, text=f"({len(scenarios)})", text_color="gray").pack(side="left")

        # CONTENT
        self.content = ctk.CTkFrame(self, fg_color="transparent")
        self.content.pack(fill="x", padx=(40, 0))

        for s in scenarios:
            row = ctk.CTkFrame(self.content, fg_color="transparent")
            row.pack(fill="x", pady=2)
            
            var = ctk.BooleanVar(value=True)
            txt = f"{s['ID']}"
            if s['Config']: txt += f" [{s['Config']}]"
            
            # Checkbox a sinistra
            chk = ctk.CTkCheckBox(row, text=txt, variable=var, height=20, font=("Arial", 12))
            chk.pack(side="left", padx=5)
            
            # Entry a destra (Nome File)
            default_out = s['Output']
            if not default_out: 
                default_out = f"{s['ID']}_{s['Config']}".replace(" ","").replace("-","_")
                
            entry = ctk.CTkEntry(row, height=24, font=("Consolas", 11))
            entry.insert(0, default_out)
            entry.pack(side="right", padx=10, fill="x", expand=True)
            
            # Se è Motor Starting, aggiungi campo "Steps"
            entry_steps = None
            if 'motor' in s['Mode'].lower():
                ctk.CTkLabel(row, text="Steps:", font=("Arial", 10)).pack(side="right", padx=2)
                entry_steps = ctk.CTkEntry(row, width=35, height=24)
                entry_steps.insert(0, "3") # Default 3 steps
                entry_steps.pack(side="right", padx=5)

            # Icona documento
            ctk.CTkLabel(row, text="📄", width=20, text_color="gray").pack(side="right")
            
            self.child_checkboxes.append({'data': s, 'var': var, 'entry': entry, 'steps': entry_steps})

    def toggle(self):
        if self.is_expanded:
            self.content.pack_forget()
            self.btn_toggle.configure(text="▶")
            self.is_expanded = False
        else:
            self.content.pack(fill="x", padx=(40, 0))
            self.btn_toggle.configure(text="▼")
            self.is_expanded = True

    def on_master_click(self):
        state = self.var_master.get()
        for child in self.child_checkboxes: child['var'].set(state)

    def get_selected(self):
        selected = []
        for c in self.child_checkboxes:
            if c['var'].get():
                # Crea copia e aggiorna Output con valore Entry
                item = c['data'].copy()
                item['Output'] = c['entry'].get().strip()
                # Recupera steps se esiste, altrimenti 1
                if c['steps']:
                    try: item['Steps'] = int(c['steps'].get())
                    except: item['Steps'] = 3
                else:
                    item['Steps'] = 1
                selected.append(item)
        return selected

class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("ETAP SLD Batch Printer")
        self.geometry("1100x800") # Allargata un po'
        
        self.runner = ETAPFinalRunner(BASE_ADDRESS)
        self.groups = []
        self.worker = None # Riferimento al thread

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1) # App espande riga 1 (main)

        # Header
        h = ctk.CTkFrame(self)
        h.grid(row=0, column=0, padx=20, pady=10, sticky="ew")
        ctk.CTkLabel(h, text="ETAP SLD Batch Printer", font=("Arial", 20, "bold")).pack(side="left", padx=20)
        
        # Settings button
        ctk.CTkButton(h, text="⚙️", width=40, command=lambda: SettingsDialog(self), 
                     fg_color="#555", hover_color="#777").pack(side="right", padx=5)
        
        ctk.CTkButton(h, text="1. Connetti", command=self.connect, fg_color="green").pack(side="right", padx=5)

        # Main
        main = ctk.CTkFrame(self)
        main.grid(row=1, column=0, padx=20, pady=10, sticky="nsew")
        main.columnconfigure(0, weight=1)
        main.columnconfigure(1, weight=1)
        main.rowconfigure(0, weight=1) # FIX: Espande il contenuto (Tree + Log) per tutta l'altezza
        
        self.tree_container = ctk.CTkScrollableFrame(main, label_text="Scenari")
        self.tree_container.grid(row=0, column=0, padx=5, pady=5, sticky="nsew")

        self.log_box = ctk.CTkTextbox(main, font=("Consolas", 12))
        self.log_box.grid(row=0, column=1, padx=5, pady=5, sticky="nsew")

        # Footer
        f = ctk.CTkFrame(self)
        f.grid(row=2, column=0, padx=20, pady=20, sticky="ew")
        
        self.btn_run = ctk.CTkButton(f, text="2. AVVIA SELEZIONATI", command=self.start, 
                                     fg_color="#D32F2F", font=("Arial", 14, "bold"), width=200, height=40)
        self.btn_run.pack(side="left", padx=20, pady=10)

        # Stop Button (inizialmente disabilitato)
        self.btn_stop = ctk.CTkButton(f, text="🛑 STOP", command=self.stop_process, 
                                      fg_color="gray", state="disabled", width=100, height=40)
        self.btn_stop.pack(side="left", padx=10, pady=10)

    def log(self, msg):
        self.log_box.insert("end", msg + "\n")
        self.log_box.see("end")

    def connect(self):
        self.log("Connessione...")
        if self.runner.connect():
            self.log("✓ Connesso")
            self.load()
        else: self.log("✗ Erore Connessione")

    def load(self):
        for w in self.tree_container.winfo_children(): w.destroy()
        self.groups = []
        scenarios = self.runner.get_scenarios()
        if not scenarios: return
        self.log(f"✓ {len(scenarios)} Scenari.")
        
        grouped = defaultdict(list)
        for s in scenarios: grouped[s['Mode'] or "Unknown"].append(s)
        
        for mode in sorted(grouped.keys()):
            grp = CollapsibleGroup(self.tree_container, title=mode.upper(), scenarios=grouped[mode])
            self.groups.append(grp)

    def start(self):
        all_selected = []
        for grp in self.groups:
            all_selected.extend(grp.get_selected())
        
        if not all_selected:
            self.log("⚠️ Seleziona almeno uno scenario.")
            return

        # --- PRE-CHECK FILES ESISTENTI ---
        out_dir = CONFIG.get("OUTPUT_DIR", os.getcwd())
        existing_conflicts = []
        
        for s in all_selected:
            base = s.get('Output', f"{s['ID']}_{s['Config']}".replace(" ","").replace("-","_"))
            
            # Calcola quali file verrebbero generati
            filenames = []
            if 'motor' in s['Mode'].lower():
                steps = s.get('Steps', 3)
                for i in range(steps):
                    filenames.append(f"{base}_T{i:02d}.pdf")
            else:
                filenames.append(f"{base}.pdf")
            
            for f in filenames:
                if os.path.exists(os.path.join(out_dir, f)):
                    existing_conflicts.append(f)

        auto_overwrite = False
        if existing_conflicts:
            count = len(existing_conflicts)
            list_str = "\n".join(existing_conflicts[:5])
            if count > 5: list_str += f"\n...e altri {count-5} file."
            
            msg = f"Attenzione! Trovati {count} file già esistenti:\n\n{list_str}\n\nVuoi sovrascriverli TUTTI automaticamente?"
            # MB_YESNO | MB_ICONWARNING | MB_TOPMOST
            resp = ctypes.windll.user32.MessageBoxW(0, msg, "Conferma Sovrascrittura Globale", 0x04 | 0x30 | 0x40000)
            
            if resp == 6: # YES
                auto_overwrite = True
                self.log(f"⚠️ Approvata sovrascrittura di {count} file.")
            else: # NO
                self.log("⛔ Annullato dall'utente (file esistenti).")
                return

        self.btn_run.configure(state="disabled")
        self.btn_stop.configure(state="normal", fg_color="#FF9800")
        
        self.worker = ETAPRunnerThread(all_selected, self.log, self.runner, auto_overwrite)
        self.worker.start()
        self.after(1000, self.check)

    def stop_process(self):
        if self.worker and self.worker.is_alive():
            self.log("\n🛑 RICHIESTA STOP INVIATA...")
            self.worker.stop()
            self.btn_stop.configure(state="disabled", text="Stopping...")

    def check(self):
        if self.worker and self.worker.is_alive():
            self.after(1000, self.check)
        else:
            # Reset UI fine processo
            self.btn_run.configure(state="normal")
            self.btn_stop.configure(state="disabled", fg_color="gray", text="🛑 STOP")
            self.log("--- FINE ---")
            
            # Porta l'app in primo piano (metodo FORTE)
            focus_app_window()

class ETAPRunnerThread(threading.Thread):
    def __init__(self, scenarios, callback, runner, auto_overwrite=False):
        super().__init__()
        self.scenarios = scenarios; self.log = callback; self.runner = runner
        self.auto_overwrite = auto_overwrite
        self.running = True

    def stop(self):
        self.running = False

    def run(self):
        self.log(f"\n🚀 START SEQUENZA ({len(self.scenarios)} items)...")
        # Focus qui
        self.log("🔍 Cerco finestra ETAP...")
        focus_etap_window()

        self.log("⌨️  MANI LONTANO DALLA TASTIERA!\n")
        time.sleep(3)
        
        for i, s in enumerate(self.scenarios, 1):
            if not self.running: 
                self.log("⛔ INTERROTTO DALL'UTENTE.")
                break

            self.log(f"\n[{i}] {s['ID']}")
            focus_etap_window()
            
            if not self.runner.select_scenario(s): 
                self.log("  ❌ Sel. Error")
                continue

            if not self.running: break 

            self.runner.run_scenario(s['ID'])
            
            for _ in range(3): 
                if not self.running: break
                time.sleep(1.0)
            if not self.running: break

            # Usa il nome file definito nella GUI
            base = s.get('Output', f"{s['ID']}_{s['Config']}".replace(" ","").replace("-","_"))
            
            if 'motor' in s['Mode'].lower():
                steps = s.get('Steps', 3)
                self.log(f"  ⚙️ Motor Mode ({steps} Steps) -> {base}...")
                
                for step_idx in range(steps):
                    if self.check_stop(): break
                    
                    filename = f"{base}_T{step_idx:02d}"
                    
                    if step_idx == 0:
                        self.log(f"      • Step 0 -> {filename}")
                        self.runner.perform_motor_step_1(filename, self.auto_overwrite)
                    else:
                        self.log(f"      • Step {step_idx} -> {filename}")
                        self.runner.perform_motor_step_advance(step_idx+1, filename, self.auto_overwrite)
            else:
                self.log(f"  ⚙️ Standard Mode -> {base}")
                self.runner.print_via_ui(base, self.auto_overwrite)
            
            self.log("  ✅ OK")
            time.sleep(1.0)

    def check_stop(self):
        if not self.running:
            self.log("⛔ Stop rilevato.")
            return True
        return False

class ETAPFinalRunner:
    def __init__(self, base): self.base = base
    def connect(self):
        try: self.conn = etap.api.connect(self.base); return True
        except: return False
    def get_scenarios(self):
        try:
            xml_resp = self.conn.scenario.getxml()
            if hasattr(xml_resp, 'strip') and (xml_resp.strip().startswith('{') or xml_resp.strip().startswith('[')):
                 try:
                    js = json.loads(xml_resp)
                    if isinstance(js, dict):
                        for key in ['Value', 'Result', 'XML', 'Data', 'd']:
                            if key in js:
                                xml_resp = js[key]; break
                 except: pass
            try: root = ET.fromstring(xml_resp)
            except: return []
            res = []
            for s in root.findall('Scenario'):
                 res.append({
                    'ID': s.get('ID', 'Unknown'), 'Mode': s.get('Mode', 'General'),
                    'Config': s.get('Config', ''), 'Output': s.get('Output', 'Untitled'),
                    'System': s.get('System', ''), 'Presentation': s.get('Presentation', ''),
                    'Revision': s.get('Revision', ''), 'StudyCase': s.get('StudyCase', ''),
                    'Executable': s.get('Executable', 'Yes')
                })
            return res
        except: return []
    def select_scenario(self, s):
        try:
            self.conn.scenario.createscenario(s['ID'], s['System'], s['Presentation'], s['Revision'], s['Config'], s['Mode'], s['Mode'].split()[-1], s['StudyCase'], '')
            return True
        except: return False
    def run_scenario(self, i): 
        try: self.conn.scenario.run(i, False)
        except: pass
    def print_via_ui(self, filename, auto_overwrite=False):
        # 1. Preparazione Path Completo
        out_dir = CONFIG.get("OUTPUT_DIR", os.getcwd())
        # Normalizza path per Windows (converte tutto in backslash) per evitare errori ETAP
        full_path = os.path.normpath(os.path.join(out_dir, filename))
        full_path_pdf = full_path + ".pdf"
        
        # 2. Controllo Interattivo Esistenza File
        if os.path.exists(full_path_pdf):
            do_delete = False
            
            if auto_overwrite:
                do_delete = True
            else:
                # Chiedi utente
                msg = f"Il file PDF esiste già:\n{filename}\n\nVuoi sovrascriverlo?"
                flags = 0x03 | 0x30 | 0x40000
                response = ctypes.windll.user32.MessageBoxW(0, msg, "Conferma", flags)
                
                if response == 6: do_delete = True # YES
                elif response == 7: # NO
                    print(f"    ⏭️ Saltato: {filename}")
                    return 
                else: return # CANCEL
            
            if do_delete:
                try: 
                    os.remove(full_path_pdf)
                    print(f"    ♻️ Cancellato: {filename}")
                    time.sleep(0.5)
                except: pass

        # 3. Automazione Stampa
        print(f"    🖨️ Stampa su: {full_path}")
        focus_etap_window() 
        pyautogui.hotkey('ctrl','p'); time.sleep(CONFIG["DELAY_PRINT_DIALOG"])
        for _ in range(7): pyautogui.press('tab'); time.sleep(0.05)
        pyautogui.press('enter'); time.sleep(CONFIG["DELAY_SAVE_DIALOG"])
        pyautogui.write(full_path); time.sleep(0.5); pyautogui.press('enter'); time.sleep(1.5)
        
    def perform_motor_step_1(self, f, auto_overwrite=False):
        pyautogui.hotkey('ctrl','tab'); time.sleep(CONFIG["DELAY_KEY_PRESS"])
        pyautogui.hotkey('ctrl','tab'); time.sleep(CONFIG["DELAY_KEY_PRESS"])
        pyautogui.press('esc'); time.sleep(CONFIG['DELAY_BETWEEN_SEQS'])
        self.print_via_ui(f, auto_overwrite)
    def perform_motor_step_advance(self, n, f, auto_overwrite=False):
        pyautogui.hotkey('ctrl','tab'); time.sleep(CONFIG["DELAY_KEY_PRESS"])
        pyautogui.hotkey('ctrl','tab'); time.sleep(CONFIG["DELAY_KEY_PRESS"])
        for _ in range(5): pyautogui.press('tab'); time.sleep(0.05)
        pyautogui.press('space'); time.sleep(CONFIG["DELAY_AFTER_SPACE"])
        pyautogui.press('esc'); time.sleep(CONFIG['DELAY_BETWEEN_SEQS'])
        self.print_via_ui(f, auto_overwrite)

if __name__ == "__main__": App().mainloop()
