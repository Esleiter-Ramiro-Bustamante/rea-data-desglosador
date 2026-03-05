"""
seguridad.py — Configuración de privacidad, auditoría y advertencias
ReaDesF1.8
"""

import os
import hashlib
from datetime import datetime


class ConfiguracionSeguridad:
    MODO_ANONIMIZAR                = False
    CREAR_LOG_AUDITORIA            = True
    LOG_DIRECTORY                  = "logs_auditoria"
    MOSTRAR_ADVERTENCIA_PRIVACIDAD = True

    @staticmethod
    def mostrar_advertencia_inicial():
        if not ConfiguracionSeguridad.MOSTRAR_ADVERTENCIA_PRIVACIDAD:
            return
        print("\n" + "=" * 70)
        print("⚠️  ADVERTENCIA DE PRIVACIDAD Y SEGURIDAD")
        print("=" * 70)
        print("  • Facturas procesadas LOCALMENTE — sin envío a internet")
        print("  • Cumple con LFPDPPP")
        if ConfiguracionSeguridad.MODO_ANONIMIZAR:
            print("  ✅ MODO ANONIMIZACIÓN ACTIVADO")
        if ConfiguracionSeguridad.CREAR_LOG_AUDITORIA:
            print(f"  📝 Log de auditoría: {ConfiguracionSeguridad.LOG_DIRECTORY}/")
        print("=" * 70)
        r = input("¿Deseas continuar? (SI/NO): ").strip().upper()
        if r not in ['SI', 'S', 'YES', 'Y']:
            print("❌ Proceso cancelado.")
            raise SystemExit(0)
        print()


class LogAuditoria:

    def __init__(self):
        self.entries = []
        if ConfiguracionSeguridad.CREAR_LOG_AUDITORIA:
            os.makedirs(ConfiguracionSeguridad.LOG_DIRECTORY, exist_ok=True)

    def _hash(self, ruta: str) -> str:
        try:
            with open(ruta, 'rb') as f:
                return hashlib.sha256(f.read()).hexdigest()
        except:
            return "NO_DISPONIBLE"

    def registrar_inicio(self, archivo: str, motor: str):
        self.entries.append({
            'ts': datetime.now(), 'evento': 'INICIO',
            'archivo': archivo,
            'hash': self._hash(archivo),
            'motor': motor
        })

    def registrar_fin(self, salida: str, filas: int, tiempo: float, motor: str):
        self.entries.append({
            'ts': datetime.now(), 'evento': 'FIN',
            'salida': salida, 'filas': filas,
            'tiempo_s': round(tiempo, 3),
            'motor': motor
        })

    def registrar_error(self, error):
        self.entries.append({
            'ts': datetime.now(),
            'evento': 'ERROR',
            'detalle': str(error)
        })

    def guardar_log(self):
        if not ConfiguracionSeguridad.CREAR_LOG_AUDITORIA:
            return
        ts   = datetime.now().strftime("%Y%m%d_%H%M%S")
        path = os.path.join(
            ConfiguracionSeguridad.LOG_DIRECTORY,
            f"auditoria_{ts}.log"
        )
        try:
            with open(path, 'w', encoding='utf-8') as f:
                f.write("=" * 70 + "\n")
                f.write("LOG DE AUDITORÍA — ReaDesF1.8\n")
                f.write("=" * 70 + "\n\n")
                for e in self.entries:
                    f.write(f"[{e['ts'].strftime('%Y-%m-%d %H:%M:%S')}]"
                            f" {e['evento']}\n")
                    for k, v in e.items():
                        if k not in ['ts', 'evento']:
                            f.write(f"  {k}: {v}\n")
                    f.write("\n")
                f.write("=" * 70 + "\nFIN DEL LOG\n" + "=" * 70 + "\n")
            print(f"  📝 Log guardado: {path}")
        except Exception as ex:
            print(f"  ⚠️  Log no guardado: {ex}")
