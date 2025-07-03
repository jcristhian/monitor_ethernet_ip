# -*- coding: utf-8 -*-
"""
Created on Thu Jul  3 14:43:46 2025

@author: Admin
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.animation import FuncAnimation
import pandas as pd
from datetime import datetime
import threading
import queue
import time
import socket
import struct
from pycomm3 import LogixDriver
import numpy as np

class PLCMonitor:
    def __init__(self, root):
        self.root = root
        self.root.title("Monitor PLC Micro 800 - Ethernet/IP")
        self.root.geometry("1200x800")
        
        # Variables de control
        self.plc = None
        self.is_connected = False
        self.is_monitoring = False
        self.data_queue = queue.Queue()
        self.data_buffer = []
        self.max_points = 100
        
        # Datos para graficar
        self.timestamps = []
        self.variable1_data = []
        self.variable2_data = []
        
        self.setup_ui()
        self.setup_plot()
        
    def setup_ui(self):
        # Frame principal
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configuración de la ventana
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(2, weight=1)
        
        # Frame de conexión
        connection_frame = ttk.LabelFrame(main_frame, text="Configuración de Conexión", padding="10")
        connection_frame.grid(row=0, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        
        # Parámetros de conexión
        ttk.Label(connection_frame, text="IP del PLC:").grid(row=0, column=0, sticky=tk.W, padx=(0, 5))
        self.ip_entry = ttk.Entry(connection_frame, width=15)
        self.ip_entry.insert(0, "192.168.101.60")
        self.ip_entry.grid(row=0, column=1, padx=(0, 10))
        
        ttk.Label(connection_frame, text="Puerto:").grid(row=0, column=2, sticky=tk.W, padx=(0, 5))
        self.port_entry = ttk.Entry(connection_frame, width=10)
        self.port_entry.insert(0, "44818")
        self.port_entry.grid(row=0, column=3, padx=(0, 10))
        
        ttk.Label(connection_frame, text="Timeout (s):").grid(row=0, column=4, sticky=tk.W, padx=(0, 5))
        self.timeout_entry = ttk.Entry(connection_frame, width=10)
        self.timeout_entry.insert(0, "5")
        self.timeout_entry.grid(row=0, column=5, padx=(0, 10))
        
        # Botones de conexión
        self.connect_btn = ttk.Button(connection_frame, text="Conectar", command=self.connect_plc)
        self.connect_btn.grid(row=0, column=6, padx=(10, 5))
        
        self.disconnect_btn = ttk.Button(connection_frame, text="Desconectar", command=self.disconnect_plc, state=tk.DISABLED)
        self.disconnect_btn.grid(row=0, column=7, padx=(5, 0))
        
        # Frame de variables
        variables_frame = ttk.LabelFrame(main_frame, text="Variables a Monitorear", padding="10")
        variables_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        
        # Variable 1
        ttk.Label(variables_frame, text="Variable 1:").grid(row=0, column=0, sticky=tk.W, padx=(0, 5))
        self.var1_entry = ttk.Entry(variables_frame, width=20)
        self.var1_entry.insert(0, "FLOAT_IN_1")
        self.var1_entry.grid(row=0, column=1, padx=(0, 10))
        
        ttk.Label(variables_frame, text="Etiqueta:").grid(row=0, column=2, sticky=tk.W, padx=(0, 5))
        self.var1_label_entry = ttk.Entry(variables_frame, width=15)
        self.var1_label_entry.insert(0, "SET POINT")
        self.var1_label_entry.grid(row=0, column=3, padx=(0, 10))
        
        # Variable 2
        ttk.Label(variables_frame, text="Variable 2:").grid(row=1, column=0, sticky=tk.W, padx=(0, 5), pady=(5, 0))
        self.var2_entry = ttk.Entry(variables_frame, width=20)
        self.var2_entry.insert(0, "FLOAT_IN_2")
        self.var2_entry.grid(row=1, column=1, padx=(0, 10), pady=(5, 0))
        
        ttk.Label(variables_frame, text="Etiqueta:").grid(row=1, column=2, sticky=tk.W, padx=(0, 5), pady=(5, 0))
        self.var2_label_entry = ttk.Entry(variables_frame, width=15)
        self.var2_label_entry.insert(0, "NIVEL DE TANQUE")
        self.var2_label_entry.grid(row=1, column=3, padx=(0, 10), pady=(5, 0))
        
        # Configuración de muestreo
        ttk.Label(variables_frame, text="Intervalo (ms):").grid(row=0, column=4, sticky=tk.W, padx=(10, 5))
        self.interval_entry = ttk.Entry(variables_frame, width=10)
        self.interval_entry.insert(0, "1000")
        self.interval_entry.grid(row=0, column=5, padx=(0, 10))
        
        ttk.Label(variables_frame, text="Máx. puntos:").grid(row=1, column=4, sticky=tk.W, padx=(10, 5), pady=(5, 0))
        self.max_points_entry = ttk.Entry(variables_frame, width=10)
        self.max_points_entry.insert(0, "100")
        self.max_points_entry.grid(row=1, column=5, padx=(0, 10), pady=(5, 0))
        
        # Botones de control
        control_frame = ttk.Frame(variables_frame)
        control_frame.grid(row=0, column=6, rowspan=2, padx=(20, 0))
        
        self.start_btn = ttk.Button(control_frame, text="Iniciar Monitoreo", command=self.start_monitoring, state=tk.DISABLED)
        self.start_btn.pack(pady=(0, 5))
        
        self.stop_btn = ttk.Button(control_frame, text="Detener Monitoreo", command=self.stop_monitoring, state=tk.DISABLED)
        self.stop_btn.pack(pady=(0, 5))
        
        self.clear_btn = ttk.Button(control_frame, text="Limpiar Datos", command=self.clear_data)
        self.clear_btn.pack(pady=(0, 5))
        
        self.export_btn = ttk.Button(control_frame, text="Exportar a Excel", command=self.export_to_excel)
        self.export_btn.pack()
        
        # Frame del gráfico
        self.plot_frame = ttk.LabelFrame(main_frame, text="Gráfico en Tiempo Real", padding="10")
        self.plot_frame.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Status bar
        self.status_var = tk.StringVar()
        self.status_var.set("Desconectado")
        status_bar = ttk.Label(main_frame, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W)
        status_bar.grid(row=3, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(10, 0))
        
    def setup_plot(self):
        # Configurar matplotlib
        plt.style.use('seaborn-v0_8')
        self.fig, self.ax = plt.subplots(figsize=(10, 6))
        self.ax.set_title("Monitor PLC Micro 800 - Tiempo Real")
        self.ax.set_xlabel("Tiempo")
        self.ax.set_ylabel("Valor")
        self.ax.grid(True, alpha=0.3)
        
        # Líneas del gráfico
        self.line1, = self.ax.plot([], [], 'b-', linewidth=0.5, label='Variable 1')
        self.line2, = self.ax.plot([], [], 'r-', linewidth=0.5, label='Variable 2')
        self.ax.legend()
        
        # Integrar matplotlib con tkinter
        self.canvas = FigureCanvasTkAgg(self.fig, master=self.plot_frame)
        self.canvas.draw()
        self.canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)
        
    def connect_plc(self):
        try:
            ip = self.ip_entry.get().strip()
            port = int(self.port_entry.get())
            timeout = float(self.timeout_entry.get())
            
            if not ip:
                messagebox.showerror("Error", "La dirección IP es requerida")
                return
            
            # Crear conexión con pycomm3
            self.plc = LogixDriver(ip + ':' + str(port))
            self.plc.open()
            
            # Verificar conexión
            if self.plc.connected:
                self.is_connected = True
                self.status_var.set(f"Conectado a {ip}:{port}")
                
                # Habilitar/deshabilitar botones
                self.connect_btn.config(state=tk.DISABLED)
                self.disconnect_btn.config(state=tk.NORMAL)
                self.start_btn.config(state=tk.NORMAL)
                
                # Deshabilitar campos de conexión
                self.ip_entry.config(state=tk.DISABLED)
                self.port_entry.config(state=tk.DISABLED)
                self.timeout_entry.config(state=tk.DISABLED)
                
                messagebox.showinfo("Éxito", "Conexión establecida correctamente")
            else:
                raise Exception("No se pudo establecer la conexión")
                
        except Exception as e:
            messagebox.showerror("Error de Conexión", f"Error al conectar con el PLC:\n{str(e)}")
            self.status_var.set("Error de conexión")
            
    def disconnect_plc(self):
        try:
            if self.is_monitoring:
                self.stop_monitoring()
            
            if self.plc:
                self.plc.close()
                self.plc = None
            
            self.is_connected = False
            self.status_var.set("Desconectado")
            
            # Habilitar/deshabilitar botones
            self.connect_btn.config(state=tk.NORMAL)
            self.disconnect_btn.config(state=tk.DISABLED)
            self.start_btn.config(state=tk.DISABLED)
            self.stop_btn.config(state=tk.DISABLED)
            
            # Habilitar campos de conexión
            self.ip_entry.config(state=tk.NORMAL)
            self.port_entry.config(state=tk.NORMAL)
            self.timeout_entry.config(state=tk.NORMAL)
            
        except Exception as e:
            messagebox.showerror("Error", f"Error al desconectar:\n{str(e)}")
            
    def start_monitoring(self):
        if not self.is_connected:
            messagebox.showerror("Error", "No hay conexión con el PLC")
            return
            
        try:
            # Obtener configuración
            self.max_points = int(self.max_points_entry.get())
            interval = int(self.interval_entry.get()) / 1000.0  # Convertir a segundos
            
            var1_name = self.var1_entry.get().strip()
            var2_name = self.var2_entry.get().strip()
            
            if not var1_name or not var2_name:
                messagebox.showerror("Error", "Ambas variables son requeridas")
                return
            
            # Iniciar monitoreo
            self.is_monitoring = True
            self.status_var.set("Monitoreando...")
            
            # Configurar botones
            self.start_btn.config(state=tk.DISABLED)
            self.stop_btn.config(state=tk.NORMAL)
            
            # Iniciar hilo de lectura
            self.monitoring_thread = threading.Thread(
                target=self.monitor_variables,
                args=(var1_name, var2_name, interval),
                daemon=True
            )
            self.monitoring_thread.start()
            
            # Iniciar animación del gráfico
            self.animation = FuncAnimation(
                self.fig, self.update_plot, interval=max(100, int(interval*1000)), 
                blit=False, repeat=True
            )
            
        except Exception as e:
            messagebox.showerror("Error", f"Error al iniciar monitoreo:\n{str(e)}")
            
    def stop_monitoring(self):
        self.is_monitoring = False
        self.status_var.set("Deteniendo monitoreo...")
        
        # Detener animación
        if hasattr(self, 'animation'):
            self.animation.event_source.stop()
            
        # Configurar botones
        self.start_btn.config(state=tk.NORMAL)
        self.stop_btn.config(state=tk.DISABLED)
        
        self.status_var.set("Monitoreo detenido")
        
    def monitor_variables(self, var1_name, var2_name, interval):
        while self.is_monitoring:
            try:
                # Leer variables del PLC
                var1_value = self.plc.read(var1_name)
                var2_value = self.plc.read(var2_name)
                
                if var1_value.error is None and var2_value.error is None:
                    timestamp = datetime.now()
                    
                    # Agregar datos al buffer
                    data_point = {
                        'timestamp': timestamp,
                        'variable1': float(var1_value.value),
                        'variable2': float(var2_value.value)
                    }
                    
                    self.data_queue.put(data_point)
                    
                else:
                    error_msg = f"Error leyendo variables: {var1_value.error or var2_value.error}"
                    self.root.after(0, lambda: self.status_var.set(error_msg))
                    
            except Exception as e:
                error_msg = f"Error en monitoreo: {str(e)}"
                self.root.after(0, lambda: self.status_var.set(error_msg))
                
            time.sleep(interval)
            
    def update_plot(self, frame):
        # Procesar datos en cola
        while not self.data_queue.empty():
            try:
                data_point = self.data_queue.get_nowait()
                self.data_buffer.append(data_point)
                
                # Mantener solo los últimos N puntos
                if len(self.data_buffer) > self.max_points:
                    self.data_buffer.pop(0)
                    
            except queue.Empty:
                break
                
        if not self.data_buffer:
            return
            
        # Preparar datos para el gráfico
        timestamps = [point['timestamp'] for point in self.data_buffer]
        var1_values = [point['variable1'] for point in self.data_buffer]
        var2_values = [point['variable2'] for point in self.data_buffer]
        
        # Actualizar líneas
        self.line1.set_data(range(len(timestamps)), var1_values)
        self.line2.set_data(range(len(timestamps)), var2_values)
        
        # Actualizar etiquetas
        var1_label = self.var1_label_entry.get() or "Variable 1"
        var2_label = self.var2_label_entry.get() or "Variable 2"
        
        self.line1.set_label(f"{var1_label}: {var1_values[-1]:.2f}")
        self.line2.set_label(f"{var2_label}: {var2_values[-1]:.2f}")
        
        # Ajustar ejes
        if len(timestamps) > 0:
            self.ax.set_xlim(0, len(timestamps)-1)
            
            all_values = var1_values + var2_values
            if all_values:
                y_min, y_max = min(all_values), max(all_values)
                margin = (y_max - y_min) * 0.1
                self.ax.set_ylim(y_min - margin, y_max + margin)
                
        self.ax.legend()
        self.canvas.draw()
        
    def clear_data(self):
        self.data_buffer.clear()
        self.timestamps.clear()
        self.variable1_data.clear()
        self.variable2_data.clear()
        
        # Limpiar gráfico
        self.line1.set_data([], [])
        self.line2.set_data([], [])
        self.ax.set_xlim(0, 1)
        self.ax.set_ylim(0, 1)
        self.canvas.draw()
        
        self.status_var.set("Datos limpiados")
        
    def export_to_excel(self):
        if not self.data_buffer:
            messagebox.showwarning("Advertencia", "No hay datos para exportar")
            return
            
        try:
            # Solicitar archivo de destino
            file_path = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
                title="Guardar datos como..."
            )
            
            if not file_path:
                return
                
            # Crear DataFrame
            df_data = []
            for point in self.data_buffer:
                df_data.append({
                    'Timestamp': point['timestamp'],
                    self.var1_label_entry.get() or 'Variable1': point['variable1'],
                    self.var2_label_entry.get() or 'Variable2': point['variable2']
                })
                
            df = pd.DataFrame(df_data)
            
            # Exportar a Excel
            with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
                df.to_excel(writer, sheet_name='Datos', index=False)
                
                # Agregar información adicional
                info_data = {
                    'Parámetro': ['IP del PLC', 'Puerto', 'Variable 1', 'Variable 2', 'Intervalo (ms)', 'Fecha de Exportación'],
                    'Valor': [
                        self.ip_entry.get(),
                        self.port_entry.get(),
                        self.var1_entry.get(),
                        self.var2_entry.get(),
                        self.interval_entry.get(),
                        datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                    ]
                }
                
                info_df = pd.DataFrame(info_data)
                info_df.to_excel(writer, sheet_name='Configuración', index=False)
                
            messagebox.showinfo("Éxito", f"Datos exportados correctamente a:\n{file_path}")
            self.status_var.set(f"Datos exportados: {len(df)} registros")
            
        except Exception as e:
            messagebox.showerror("Error", f"Error al exportar datos:\n{str(e)}")

def main():
    root = tk.Tk()
    app = PLCMonitor(root)
    root.mainloop()

if __name__ == "__main__":
    main()
