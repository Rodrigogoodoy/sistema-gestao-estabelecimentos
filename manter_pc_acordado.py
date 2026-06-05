"""
Impede o Windows de suspender o PC enquanto os workers rodam.
Deixe essa janela aberta. Feche apenas quando terminar o processamento.
"""
import ctypes
import time

ES_CONTINUOUS       = 0x80000000
ES_SYSTEM_REQUIRED  = 0x00000001
ES_DISPLAY_REQUIRED = 0x00000002

ctypes.windll.kernel32.SetThreadExecutionState(
    ES_CONTINUOUS | ES_SYSTEM_REQUIRED | ES_DISPLAY_REQUIRED
)

print("=" * 48)
print("  PC configurado para NAO suspender.")
print("  Mantenha esta janela aberta.")
print("  Feche somente apos o processamento terminar.")
print("=" * 48)

try:
    while True:
        time.sleep(60)
except KeyboardInterrupt:
    ctypes.windll.kernel32.SetThreadExecutionState(ES_CONTINUOUS)
    print("\nModo normal restaurado.")
