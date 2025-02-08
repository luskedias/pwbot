import tkinter as tk
from tkinter import CENTER, ttk, messagebox
import pyautogui
import pytesseract
from PIL import Image, ImageTk, ImageOps
import time
import re
import win32gui
import pygetwindow as gw
import threading
import json
import logging
import concurrent.futures
import mss
import keyboard # Import da biblioteca keyboard

# Configuração do Tesseract OCR
pytesseract.pytesseract.tesseract_cmd = r'.\Tesseract-OCR\tesseract.exe'

# Configuração básica do logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')


class Application(tk.Tk):
    def __init__(self):
        super().__init__()

        self.title("Leitor de Atributos Perfect World")
        self.geometry("350x600")

        self.running = False
        self.button_position = None
        self.equip_position = None
        self.icon_region = None
        self.atributos_sets = []
        self.selected_window = None
        self.config_file = "config.json"
        self.config_data = {}

        self.canvas = None
        self.rect_start = None
        self.rect_id = None
        self.screenshot = None
        self.screenshot_tk = None
        self.icon_image = None

        self.initial_gui_hidden = False
        self.hotkey = None # Inicializa self.hotkey
        self.hotkey_listener_thread = None # Inicializa o thread listener
        

        self.create_initial_gui()
        self.setup_styles()

        self.start_hotkey_listener() # Inicia o listener de hotkey

    def start_hotkey_listener(self):
        """Inicia um thread para escutar a hotkey globalmente."""
        self.hotkey_listener_thread = threading.Thread(target=self._hotkey_listener, daemon=True) # Daemon para terminar com a app
        self.hotkey_listener_thread.start()

    def _hotkey_listener(self):
        """Função rodando no thread para detectar a hotkey."""
        while True:
            if self.hotkey: # Verifica se uma hotkey foi definida
                keyboard.wait(self.hotkey) # Aguarda a hotkey ser pressionada
                self.on_hotkey_event() # Chama a função de tratamento do evento
            else:
                time.sleep(1) # Espera um pouco se nenhuma hotkey estiver definida para não consumir CPU desnecessariamente


    def on_hotkey_event(self):
        """Função chamada quando a hotkey é pressionada."""
        if self.running:
            self.toggle_script() # Se o script estiver rodando, para o script
        else:
            self.start_program() # Se o script não estiver rodando, inicia o script


    def setup_styles(self):
        """Configura os estilos ttk para destacar widgets."""
        style = ttk.Style()

        # Estilo para botões selecionados
        style.configure("Filled.TButton",
                        background="#3498db",
                        foreground="grey",
                        relief="raised",
                        borderwidth=3)
        style.map("Filled.TButton",
                  background=[("active", "#2980b9")],
                  relief=[("pressed", "groove")])

        # Estilo para Combobox NORMAL (antes de ser preenchido)
        style.configure("TCombobox",
                        fieldbackground="white",
                        foreground="black",
                        selectbackground="#3498db",
                        selectforeground="white",
                        arrowcolor="#3498db")

        # Estilo para Combobox selecionado (Filled)
        style.configure("Filled.TCombobox",
                        fieldbackground="#2c3e50",
                        foreground="green",
                        selectbackground="#3498db",
                        selectforeground="black",
                        arrowcolor="black")
        style.map("Filled.TCombobox",
                  fieldbackground=[("readonly", "#2c3e50")],
                  foreground=[("readonly", "black")])

        # Estilo para Entry NORMAL (antes de ser preenchido)
        style.configure("TEntry",
                        fieldbackground="black",
                        foreground="black",
                        insertcolor="black",
                        selectbackground="#3498db",
                        selectforeground="black")

        # Estilo para Entry selecionado (Filled)
        style.configure("Filled.TEntry",
                        fieldbackground="#2c3e50",
                        foreground="black",
                        insertcolor="black",
                        selectbackground="#3498db",
                        selectforeground="black")

    def create_initial_gui(self):
        self.title("Macro de Gravura Perfect World")
        main_frame = ttk.Frame(self)
        main_frame.pack(pady=20, padx=20)

        button_width = 20
        button_pady = 5

        self.gravura_button = ttk.Button(main_frame, text="Setar botão GRAVURA", command=self.set_gravura_position,
                                        width=button_width)
        self.gravura_button.grid(row=0, column=0, pady=button_pady, sticky="ew")

        self.equip_button = ttk.Button(main_frame, text="Setar Equipamento", command=self.set_equip_position,
                                        width=button_width)
        self.equip_button.grid(row=1, column=0, pady=button_pady, sticky="ew")

        self.regiao_button = ttk.Button(
            main_frame,
            text="Setar região de captura",
            command=self.set_capture_region,
            width=button_width
        )
        self.regiao_button.grid(row=2, column=0, pady=button_pady, sticky="ew")

        self.add_atributo_button = ttk.Button(main_frame, text="Adicionar Conjunto de Atributos",
                                                 command=self.add_atributo_set, width=button_width)
        self.add_atributo_button.grid(row=3, column=0, pady=button_pady, sticky="ew")

        self.atributos_frame = ttk.Frame(main_frame)
        self.atributos_frame.grid(row=4, column=0, pady=5, sticky="ew")

        self.select_window_button = ttk.Button(main_frame, text="Selecionar Janela do PW", command=self.select_window,
                                                width=button_width)
        self.select_window_button.grid(row=5, column=0, pady=button_pady, sticky="ew")

        self.selected_window_label = ttk.Label(main_frame, text="Nenhuma janela selecionada")
        self.selected_window_label.grid(row=6, column=0, sticky="ew")

        # *** NOVO BOTÃO PARA SETAR HOTKEY ***
        self.hotkey_button = ttk.Button(main_frame, text="Setar Hotkey", command=self.set_hotkey,
                                            width=button_width)
        self.hotkey_button.grid(row=7, column=0, pady=button_pady, sticky="ew")

        self.hotkey_label = ttk.Label(main_frame, text="Nenhuma hotkey definida")
        self.hotkey_label.grid(row=8, column=0, sticky="ew") # Ajustado para row=8

        ttk.Button(main_frame, text="Iniciar", command=self.start_program, width=button_width).grid(row=9, column=0, # Ajustado para row=9
                                                                                                   pady=5,
                                                                                                   sticky="ew")

        main_frame.columnconfigure(0, weight=1)
        for i in range(10): # Ajustado para range(10)
            main_frame.rowconfigure(i, weight=1)

        # Label para a imagem (inicialmente vazia)
        self.image_label = ttk.Label(main_frame)
        self.image_label.grid(row=10, column=0, pady=5, sticky="ew") # Ajustado para row=10


        credito_label = ttk.Label(self, text="Crédito: Desenvolvido por JettaStage.\nSoftware gratuito.")
        credito_label.pack(side=tk.BOTTOM, pady=(0, 10))
        credito_label.config(justify=CENTER)

        # Carrega as configurações e atualiza a interface
        self.load_config()
        self.recreate_attribute_frames()
        self.update_gui_from_config()


    def update_gui_from_config(self):
        """Atualiza o estado da GUI com base nas configurações carregadas."""
        if self.config_data:
            if "button_position" in self.config_data:
                self.button_position = tuple(self.config_data["button_position"])
                self.mark_field_as_filled(self.gravura_button)
            if "equip_position" in self.config_data:
                self.equip_position = tuple(self.config_data["equip_position"])
                self.mark_field_as_filled(self.equip_button)
            if "icon_region" in self.config_data:
                self.icon_region = tuple(self.config_data["icon_region"])
                self.mark_field_as_filled(self.regiao_button)
                # Recria a imagem, se existir
                if self.screenshot:
                    try:
                        cropped_image = self.screenshot.crop(self.icon_region)
                        self.icon_image = ImageTk.PhotoImage(cropped_image)
                        self.image_label.config(image=self.icon_image)
                    except Exception as e:
                        logging.error(f"Erro ao recriar imagem da região: {e}")

            if "selected_window" in self.config_data:
                self.selected_window = self.config_data["selected_window"]
                self.selected_window_label.config(text=f"Janela selecionada: {self.selected_window}")
                self.mark_field_as_filled(self.selected_window_label)

            if "hotkey" in self.config_data: # Carrega a hotkey
                self.hotkey = self.config_data["hotkey"]
                self.hotkey_label.config(text=f"Hotkey definida: {self.hotkey}")
                self.mark_field_as_filled(self.hotkey_label)


    def update_gui(self):
        self.update()

    def mark_field_as_filled(self, widget):
        """Aplica o estilo 'Filled' ao widget."""
        if isinstance(widget, ttk.Combobox):
            widget.config(style="Filled.TCombobox")
        elif isinstance(widget, ttk.Entry):
            widget.config(style="Filled.TEntry")
        elif isinstance(widget, ttk.Label):
            widget.config(foreground="green")
        elif isinstance(widget, ttk.Button):
            widget.config(style="Filled.TButton")


    def add_atributo_set(self, data=None):
        atributos_disponiveis = [
            "Escolha o atributo", "CONSTITUICAO", "FORCA", "DESTREZA", "INTELIGENCIA",
            "ATQ FISICO", "ATQM", "NIVEL DE ATAQUE",
            "NIVEL DE DEFESA", "CRITICO", "DEFM", "DEF",
            "REDUÇÃO DE DANO FÍSICO", "REDUÇÃO DE DANO DOS CINCO ELEMENTOS"
        ]

        frame = ttk.Frame(self.atributos_frame)
        frame.pack(pady=2, fill=tk.X)

        atributo_widgets = {}

        for i in range(3):
            inner_frame = ttk.Frame(frame)
            inner_frame.pack(pady=2, fill=tk.X)

            label = ttk.Label(inner_frame, text=f"Atributo {len(self.atributos_sets) * 3 + i + 1}:")
            label.pack(side=tk.LEFT)

            combo = ttk.Combobox(inner_frame, values=atributos_disponiveis)
            combo.pack(side=tk.LEFT)
            combo.current(0)

            entry = ttk.Entry(inner_frame, width=5)
            entry.pack(side=tk.LEFT)
            entry.insert(0, "0")

            atributo_widgets[f"atributo{len(self.atributos_sets) * 3 + i + 1}"] = {"combo": combo, "entry": entry}

            if data:
                combo.current(data[i]["combo_index"])
                entry.delete(0, tk.END)
                entry.insert(0, data[i]["entry_value"])
                self.mark_field_as_filled(combo)
                self.mark_field_as_filled(entry)

        remove_button = ttk.Button(frame, text="X", command=lambda f=frame: self.remove_atributo_set(f))
        remove_button.pack(side=tk.RIGHT)

        self.atributos_sets.append(atributo_widgets)

    def remove_atributo_set(self, frame):
        for idx, atributo_set in enumerate(self.atributos_sets):
            first_widget = next(iter(atributo_set.values()))['combo']
            if first_widget.winfo_toplevel() == frame.winfo_toplevel():
                del self.atributos_sets[idx]
                frame.destroy()
                self.save_config()
                return

    def save_config(self):
        config_data = {
            "button_position": self.button_position,
            "equip_position": self.equip_position,
            "icon_region": self.icon_region,
            "selected_window": self.selected_window,
            "atributos_sets": [],
            "hotkey": self.hotkey # Salva a hotkey no config
        }

        for atributo_set in self.atributos_sets:
            set_data = []
            for widget_data in atributo_set.values():
                set_data.append({
                    "combo_index": widget_data["combo"].current(),
                    "entry_value": widget_data["entry"].get(),
                })
            config_data["atributos_sets"].append(set_data)

        try:
            with open(self.config_file, "w") as f:
                json.dump(config_data, f)
            logging.info("Configurações salvas.")
        except Exception as e:
            logging.error(f"Erro ao salvar configurações: {e}")

    def load_config(self):
        try:
            with open(self.config_file, "r") as f:
                self.config_data = json.load(f)
            logging.info("Configurações carregadas.")
            if "hotkey" in self.config_data: # Carrega a hotkey se existir
                self.hotkey = self.config_data["hotkey"]
            else:
                self.hotkey = None # Garante que self.hotkey seja inicializado mesmo se não houver no config
        except (FileNotFoundError, json.JSONDecodeError, TypeError):
            self.config_data = {}
            self.hotkey = None # Garante que self.hotkey seja inicializado em caso de erro/arquivo não encontrado
            logging.warning("Arquivo de configuração corrompido ou não encontrado. Configurações resetadas.")
        except Exception as e:
            self.config_data = {}
            self.hotkey = None # Garante que self.hotkey seja inicializado em caso de erro genérico
            logging.error(f"Erro ao carregar configurações: {e}")


    def set_hotkey(self):
        self.capture_hotkey_window = tk.Toplevel(self)
        self.capture_hotkey_window.title("Setar Hotkey")
        tk.Label(self.capture_hotkey_window, text="Pressione a hotkey desejada:").pack(pady=10, padx=10)
        self.capture_hotkey_window.focus_set()
        self.current_hotkey_stringvar = tk.StringVar()
        tk.Label(self.capture_hotkey_window, textvariable=self.current_hotkey_stringvar).pack(pady=5)

        self.countdown_label = tk.Label(self.capture_hotkey_window, text="") # Label para o timer
        self.countdown_label.pack(pady=5)

        self.countdown_seconds = 5 # Define a duração do timer

        threading.Thread(target=self._capture_hotkey_manual_thread, daemon=True).start() # *** Inicia a thread de captura MANUALMENTE AQUI, ANTES do timer ***
        self.update_countdown() # Inicia a contagem decrescente


    def update_countdown(self):
        if self.capture_hotkey_window: # *** Check if the window still exists ***
            if self.countdown_seconds > 0:
                logging.info("Configurando countdown_label com texto do timer...") # Log ANTES
                self.countdown_label.config(text=f"Você tem {self.countdown_seconds} segundos para pressionar UMA tecla...")
                self.countdown_seconds -= 1
                self.after(1000, self.update_countdown) # Agenda a próxima atualização após 1 segundo
            else:
                # *** REMOVIDO: self.countdown_label.config(text="Tempo esgotado. Se nenhuma tecla foi capturada, tente novamente.") # Informa timeout ***
                logging.info("Tempo esgotado. Janela será fechada.") # Mantém o log de timeout
                self.after(1500, self.capture_hotkey_window.destroy) # Fecha a janela após um pequeno delay (opcional)
        else:
            logging.info("Janela capture_hotkey_window já foi destruída. Interrompendo update_countdown.") # Log when window is gone.
            return # Stop the countdown if the window is destroyed


    def _capture_hotkey_manual_thread(self):
        """Thread para capturar manualmente a hotkey usando keyboard.record() por um período."""
        logging.info("Thread _capture_hotkey_manual_thread iniciada (gravação manual - paralela ao contador)") # Log de início
        self.capture_hotkey_window.focus_force() # Forçar o foco

        logging.info("Iniciando keyboard.start_recording()") # Log ANTES de start_recording
        keyboard.start_recording()
        logging.info("keyboard.start_recording() executado") # Log DEPOIS de start_recording

        time.sleep(4.9) # *** Reduzido para 4.9 segundos para terminar um pouco antes do contador UI ***

        logging.info("Iniciando keyboard.stop_recording()") # Log ANTES de stop_recording
        events = keyboard.stop_recording()
        logging.info("keyboard.stop_recording() executado") # Log DEPOIS de stop_recording
        logging.info(f"Eventos capturados (gravação manual - paralela): {events}") # Log dos eventos capturados

        if events:
            # Procura o primeiro evento 'down' na lista e usa o nome dessa tecla como hotkey
            for event in events:
                if event.event_type == keyboard.KEY_DOWN: # Verifica se é um evento de tecla pressionada
                    hotkey_str = event.name
                    self.after(0, self.save_hotkey, hotkey_str) # Chama save_hotkey na thread principal
                    return # Sai da função após encontrar o primeiro evento 'down'
            else: # Se não encontrar nenhum evento 'down' em 4.9 segundos (improvável)
                self.after(0, self.capture_hotkey_window.destroy) # Fecha a janela se nenhuma tecla for capturada
                logging.info("Nenhuma tecla hotkey capturada durante o tempo limite (gravação paralela).")
        else: # Se não houver eventos (muito improvável)
            self.after(0, self.capture_hotkey_window.destroy)
            logging.info("Nenhum evento de teclado capturado (gravação paralela).")


    def save_hotkey(self, hotkey_str): # Agora hotkey_str é passado como argumento
        logging.info(f"Função save_hotkey iniciada com hotkey_str: {hotkey_str}") # 8. Adicionar este log no início de save_hotkey
        self.hotkey = hotkey_str # Salva a hotkey
        self.hotkey_label.config(text=f"Hotkey definida: {self.hotkey}") # Atualiza o Label
        self.mark_field_as_filled(self.hotkey_label)
        self.save_config()
        self.capture_hotkey_window.destroy()
        logging.info(f"Hotkey definida e salva: {self.hotkey}") # 9. Adicionar este log no final de save_hotkey


    def set_gravura_position(self):
        self.capture_position(callback=self.save_gravura_position)

    def save_gravura_position(self, coords):
        self.button_position = coords
        self.mark_field_as_filled(self.gravura_button)
        self.save_config()
        logging.info(f"Posição da Gravura: {self.button_position}")

    def set_equip_position(self):
        self.capture_position(callback=self.save_equip_position)

    def save_equip_position(self, coords):
        self.equip_position = coords
        self.mark_field_as_filled(self.equip_button)
        self.save_config()
        logging.info(f"Posição do Equipamento: {self.equip_position}")

    def select_window(self):
        windows = gw.getAllWindows()
        window_titles = [window.title for window in windows if window.title]

        if not window_titles:
            messagebox.showinfo("Aviso", "Nenhuma janela aberta encontrada.")
            return

        select_window_window = tk.Toplevel(self)
        select_window_window.title("Selecionar Janela do PW")
        window_width = 300
        window_height = 250
        select_window_window.geometry(f"{window_width}x{window_height}")

        listbox = tk.Listbox(select_window_window)
        window_titles.sort()
        for title in window_titles:
            listbox.insert(tk.END, title)
        listbox.pack(fill=tk.BOTH, expand=True)

        def on_select(select_window_window, listbox):
            try:
                selected_title = listbox.get(listbox.curselection())
                self.selected_window = selected_title
                self.selected_window_label.config(text=f"Janela selecionada: {self.selected_window}")
                self.mark_field_as_filled(self.selected_window_label)
                select_window_window.destroy()
                self.save_config()
            except tk.TclError:
                pass
            except Exception as e:
                logging.error(f"Erro ao selecionar janela: {e}")

        select_button = ttk.Button(select_window_window, text="Selecionar",
                                        command=lambda: on_select(select_window_window, listbox))
        select_button.pack()

    def capture_position(self, callback):
        overlay = tk.Toplevel(self)
        overlay.attributes('-fullscreen', True)
        overlay.attributes('-alpha', 0.3)
        overlay.overrideredirect(True)

        def on_click(event):
            x, y = event.x, event.y
            callback((x, y))
            overlay.destroy()

        overlay.bind("<Button-1>", on_click)

    def set_capture_region(self):
        if not self.initial_gui_hidden:
            self.withdraw()
            self.initial_gui_hidden = True

        self.capture_window = tk.Toplevel(self)
        self.capture_window.attributes('-fullscreen', True)
        self.capture_window.attributes('-alpha', 0.3)
        self.capture_window.overrideredirect(True)

        self.canvas = tk.Canvas(self.capture_window, width=self.winfo_screenwidth(), height=self.winfo_screenheight(),
                                highlightthickness=0)
        self.canvas.pack(fill=tk.BOTH, expand=tk.YES)

        self.instruction_text = self.canvas.create_text(
            self.winfo_screenwidth() // 2, self.winfo_screenheight() // 2,
            text="1º Posicione o mouse sobre a ARMADURA e clique com o mouse;\n2º Pressione a tecla 'a' para tirar uma screenshot;\n3º Clique e arraste para selecionar a região de captura.",
            font=("Arial", 20, "bold"),
            fill="black",
            justify=CENTER
        )

        self.capture_window.bind("<KeyRelease-a>", self.capture_screenshot)

        self.capture_window.attributes('-topmost', True)
        self.capture_window.lift()
        self.capture_window.focus_set()

    def capture_screenshot(self, event):
        self.canvas.delete(self.instruction_text)
        # Use mss for faster screenshot
        with mss.mss() as sct:
            monitor = sct.monitors[1] # Main monitor
            self.screenshot = Image.frombytes("RGB", (monitor["width"], monitor["height"]), sct.grab(monitor).rgb)
        self.screenshot_tk = ImageTk.PhotoImage(self.screenshot)
        self.canvas.create_image(0, 0, anchor=tk.NW, image=self.screenshot_tk)

        self.canvas.bind("<Button-1>", self.start_rect)
        self.canvas.bind("<B1-Motion>", self.draw_rect)
        self.canvas.bind("<ButtonRelease-1>", self.end_rect)

        self.capture_window.config(cursor="crosshair")

    def start_rect(self, event):
        self.rect_start = (event.x, event.y)

    def draw_rect(self, event):
        if self.rect_id:
            self.canvas.delete(self.rect_id)
        x1, y1 = self.rect_start
        x2, y2 = event.x, event.y
        self.rect_id = self.canvas.create_rectangle(x1, y1, x2, y2, outline="red", width=4, dash=(4, 2))

    def end_rect(self, event):
        x1, y1 = self.rect_start
        x2, y2 = event.x, event.y

        left = min(x1, x2)
        top = min(y1, y2)
        right = max(x1, x2)
        bottom = max(y1, y2)

        width = right - left
        height = bottom - top

        self.icon_region = (left, top, width, height)

        self.capture_window.config(cursor="arrow")
        self.capture_window.destroy()

        if self.initial_gui_hidden:
            self.deiconify()
            self.initial_gui_hidden = False

        logging.info(f"Região de captura: {self.icon_region}")

        # Converter região de captura em imagem
        try:
            cropped_image = self.screenshot.crop((left, top, right, bottom))
            self.icon_image = ImageTk.PhotoImage(cropped_image)

            # Exibir imagem na GUI inicial
            if hasattr(self, 'image_label'):
                self.image_label.config(image=self.icon_image)
            else:
                self.image_label = ttk.Label(self, image=self.icon_image)
                self.image_label.pack(pady=5)
        except Exception as e:
            logging.error(f"Erro ao criar/exibir imagem da região: {e}")
            messagebox.showerror("Erro", "Erro ao capturar a região.  Certifique-se de que a janela do jogo está visível.")
            return

        self.mark_field_as_filled(self.regiao_button)
        self.save_config()

    def start_program(self):
        if not all([self.button_position, self.equip_position, self.icon_region, self.selected_window]):
            messagebox.showerror("Erro",
                                 "É necessário definir todas as posições, a região de captura e selecionar uma janela.")
            return

        self.atributos_sets_to_check = []
        for atributo_set in self.atributos_sets:
            atributos = []
            valores = []
            for i in range(3):
                key = f"atributo{self.atributos_sets.index(atributo_set) * 3 + i + 1}"
                atributo = atributo_set[key]["combo"].get()
                try:
                    valor = int(atributo_set[key]["entry"].get())
                except ValueError:
                    messagebox.showerror("Erro", f"Valor inválido para o atributo '{atributo}'. Insira um número inteiro.")
                    return

                if atributo != "Escolha o atributo":
                    atributos.append(atributo)
                    valores.append(str(valor))
            if atributos:
                self.atributos_sets_to_check.append((atributos, valores))

        logging.info(f"Conjuntos de atributos para verificar: {self.atributos_sets_to_check}")

        try:
            activate_window(self.selected_window)
        except Exception as e:
            messagebox.showerror("Erro", f"Não foi possível ativar a janela do jogo. Erro: {e}")
            return

        self.destroy_initial_gui()
        self.create_running_gui()

        self.attributes('-topmost', True)
        self.attributes('-topmost', False)
        self.attributes('-topmost', True)
        self.lift()
        self.after(100, self.focus_force)

        self.running = True
        self.script_thread = threading.Thread(target=self.run_script)
        self.script_thread.start()

    def destroy_initial_gui(self):
        for widget in self.winfo_children():
            widget.destroy()

    def create_running_gui(self):
        self.geometry("300x400")

        self.start_button = ttk.Button(self, text="Parar", command=self.toggle_script)
        self.start_button.pack(pady=5)

        self.texto_lido_label = ttk.Label(self, text="Texto Lido:")
        self.texto_lido_label.pack()

        self.text_area = tk.Text(self, height=8, width=35, wrap=tk.WORD)
        self.text_area.pack()

        self.log_text = tk.Text(self, height=8, width=35, wrap=tk.WORD)
        self.log_text.pack(pady=5)

        self.return_button = ttk.Button(self, text="Voltar ao Início.", command=self.return_to_initial_gui)
        self.return_button.pack(pady=5)

    def toggle_script(self):
        self.running = not self.running
        if self.running:
            self.start_button.config(text="Parar")
            logging.info("Iniciando o script...")
            self.script_thread = threading.Thread(target=self.run_script)
            self.script_thread.start()
        else:
            self.start_button.config(text="Iniciar")
            logging.info("Script parado.")

    def run_script(self):
        try:
            activate_window(self.selected_window)

            with concurrent.futures.ThreadPoolExecutor(max_workers=4) as executor: # Use ThreadPoolExecutor for parallel tasks. Adjust max_workers as needed.
                while self.running:
                    # Submit tasks to the executor for parallel processing
                    future_capture = executor.submit(hover_and_capture_icon, self.button_position, self.icon_region, self.equip_position)

                    screenshot = future_capture.result() # Wait for screenshot to be captured
                    future_ocr = executor.submit(extract_text_from_image, screenshot) # Submit OCR task

                    extracted_text = future_ocr.result() # Wait for OCR result

                    if extracted_text:
                        logging.info(f"Texto extraído: {extracted_text}")

                        self.text_area.delete("1.0", tk.END)
                        self.text_area.insert(tk.END, extracted_text)
                        self.text_area.see(tk.END)

                        for atributos_set, valores_set in self.atributos_sets_to_check:
                            lista_atributos_lidos = []
                            lista_valores_lidos = []

                            linhas = extracted_text.strip().split('\n')
                            for linha in linhas:
                                match = re.match(r"(.+)\s+(\d+)", linha)
                                if match:
                                    atributo_lido = match.group(1).strip()
                                    valor_lido = match.group(2).strip()
                                    lista_atributos_lidos.append(atributo_lido)
                                    lista_valores_lidos.append(valor_lido)

                            logging.info(f"Atributos lidos: {lista_atributos_lidos}")
                            logging.info(f"Valores lidos: {lista_valores_lidos}")

                            if len(atributos_set) == len(lista_atributos_lidos):
                                conjunto_usuario = set(zip(atributos_set, valores_set))
                                conjunto_lido = set(zip(lista_atributos_lidos, lista_valores_lidos))

                                if conjunto_usuario == conjunto_lido:
                                    self.log(f"Combinação de atributos encontrada: {conjunto_usuario}")
                                    self.running = False
                                    self.start_button.config(text="Iniciar")
                                    break

                    if not extracted_text:
                        time.sleep(0.05) # Reduced sleep time

        except Exception as e:
            self.log(f"Erro no script: {e}")
            logging.exception("Erro no script:")

        finally:
            self.running = False
            self.start_button.config(text="Iniciar")


    def recreate_attribute_frames(self):
        """Recria os frames de atributos, limpa a lista e adiciona a partir dos dados."""
        for widget in self.atributos_frame.winfo_children():
            widget.destroy()

        self.atributos_sets = []
        for atributo_set_data in self.config_data.get("atributos_sets", []):
            self.add_atributo_set(data=atributo_set_data)


    def return_to_initial_gui(self):
        self.running = False
        if hasattr(self, 'script_thread') and self.script_thread.is_alive():
            self.script_thread.join()

        self.destroy_running_gui()
        self.load_config()
        self.create_initial_gui()
        self.geometry("350x600")


    def destroy_running_gui(self):
        for widget in self.winfo_children():
            widget.destroy()

    def log(self, message):
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)
        logging.info(message)



# Funções auxiliares
def capture_screen(region=None):
    with mss.mss() as sct: # Use mss for faster screenshot
        if region:
            monitor = {"top": region[1], "left": region[0], "width": region[2], "height": region[3]}
            sct_img = sct.grab(monitor)
            screenshot = Image.frombytes("RGB", sct_img.size, sct_img.bgra, "raw", "BGRX") # Correctly handle BGRA to RGB
        else:
            monitor = sct.monitors[1] # Main monitor
            sct_img = sct.grab(monitor)
            screenshot = Image.frombytes("RGB", (monitor["width"], monitor["height"]), sct_img.bgra, "raw", "BGRX") # Correctly handle BGRA to RGB
        return screenshot


def preprocess_image(image):
    """Pré-processa a imagem para melhorar a qualidade do OCR."""
    gray_image = ImageOps.grayscale(image) # Converte para tons de cinza
    enhanced_image = ImageOps.autocontrast(gray_image) # Melhora o contraste automaticamente
    return enhanced_image


def extract_text_from_image(image):
    try:
        preprocessed_image = preprocess_image(image) # Pré-processa image antes do OCR
        text = pytesseract.image_to_string(preprocessed_image, config='--psm 6') # PSM 6 para bloco de texto uniforme
        return text
    except Exception as e:
        logging.error(f"Erro ao extrair texto: {e}")
        return None

def hover_and_capture_icon(icon_position, capture_region, equip_position):
    pyautogui.moveTo(icon_position)
    time.sleep(0.1)
    pyautogui.leftClick(icon_position)
    time.sleep(1)
    pyautogui.moveTo(equip_position)
    time.sleep(0.6)
    pyautogui.rightClick(equip_position)
    time.sleep(0.4)

    screenshot = capture_screen(region=capture_region)
    return screenshot

def activate_window(window_title):
    try:
        hwnd = win32gui.FindWindow(None, window_title)
        if hwnd:
            win32gui.SetForegroundWindow(hwnd)
        else:
            windows = gw.getWindowsWithTitle(window_title)
            if windows:
                windows[0].activate()
            else:
                logging.warning(f"Janela '{window_title}' não encontrada.")
        time.sleep(0.05) # Redução do tempo de espera
    except Exception as e:
        logging.error(f"Erro ao ativar a janela: {e}")


if __name__ == "__main__":
    app = Application()
    app.mainloop()
