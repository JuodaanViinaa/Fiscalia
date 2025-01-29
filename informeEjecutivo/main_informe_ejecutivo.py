from organizador import organizar
from informe_ejecutivo import informe_ejecutivo
from audiencias import audiencias
from ordenes import ordenes
from enviar_correo import enviar_mensaje
import datetime

# Este codigo realiza el informe ejecutivo completo
# Por seguridad y practicidad se recomienda hacer el analisis por pasos:
# Primero descomentar solo la fila de path y organizar y correr el código. Eso organizará los archivos en las carpetas
# correspondientes. Sin embargo, depende del nombre de cada archivo, así que no es infalible. Se recomienda revisar
# las carpetas a mano para garantizar que cada cosa este en su sitio.
# Despues comentar "organizar" y descomentar "informe ejecutivo" y correr nuevamente. Eso realiza el informe principal.
# Despues comentar "informe_ejecutivo" y descomentar "audiencias". Eso hace el reporte de audiencias.

path = f'/home/daniel/PycharmProjects/Fiscalia/informeEjecutivo/'
organizar(f'{path}/{datetime.date.today().strftime("%Y%m%d")}/')
informe_ejecutivo(f'{path}/{datetime.date.today().strftime("%Y%m%d")}/Informe diario/', f'{path}/Formato_informe.xlsx')
audiencias(f'{path}/{datetime.date.today().strftime("%Y%m%d")}/Audiencias/', f'{path}/Formato_audiencias.xlsx')
ordenes(f'{path}/{datetime.date.today().strftime("%Y%m%d")}/Ordenes/', f'{path}/Formato_ordenes.xlsx')
enviar_mensaje("maldonadodaniel96@outlook.com")
