import os
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow

# Definimos el alcance estricto de lectura y escritura en Calendar
SCOPES = ['https://www.googleapis.com/auth/calendar']

def main():
    creds = None
    # Validamos que el archivo de credenciales de la consola exista
    if not os.path.exists('credentials.json'):
        print("❌ ERROR: No se encuentra el archivo 'credentials.json' en esta carpeta.")
        print("Por favor descárgalo de Google Cloud Console antes de continuar.")
        return

    # Ejecutamos el flujo de autorización local
    flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
    creds = flow.run_local_server(port=0)
    
    # Guardamos el token de acceso para producción
    with open('token.json', 'w') as token:
        token.write(creds.to_json())
        
    print("\n✅ ¡ÉXITO ROTUNDO!")
    print("Se ha generado el archivo 'token.json' correctamente en esta carpeta.")
    print("Ya puedes eliminar de forma segura este script temporal ('generar_token.py').")

if __name__ == '__main__':
    main()