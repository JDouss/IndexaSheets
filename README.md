# IndexaSheets
Este script de Google Apps Script te permite importar automáticamente tus datos de cartera de Indexa Capital directamente a una hoja de cálculo de Google Sheets. Está inspirado en el trabajo de Inversor Provinciano, pero extrae y presenta información adicional para ofrecer una visión más completa de tus inversiones.

---

## Características

* **Importación Automática:** Consigue tus datos de Indexa Capital con un solo clic.
* **Dashboard Personalizado:** Genera un `Dashboard` con un resumen de tu cartera y un desglose detallado por fondo de inversión.
* **Formato Consistente:** Los datos se formatean automáticamente con monedas, porcentajes y reglas condicionales para una fácil lectura.
* **Seguridad:** Tus credenciales (API Token y ID de cuenta) se guardan de forma segura usando `PropertiesService`.

---

## Primeros Pasos

Sigue estos pasos detallados para configurar y usar el script en tu Google Sheet:

### 1. Obtener tu Token de la API de Indexa Capital

Antes de empezar, necesitarás solicitar un **token de la API** y tener a mano tu **ID de cuenta** de Indexa Capital.

* Accede a tu **Área de Cliente** en Indexa Capital.
* Navega a `Tu nombre` > `Configuración de usuario` > `Aplicaciones`.
* Genera y copia tu token.
* También anota tu ID de cuenta (por ejemplo, `PD123ABC`).

### 2. Crear una Nueva Hoja de Cálculo en Google Sheets

* Ve a [Google Sheets](https://docs.google.com/spreadsheets/u/0/) y crea una hoja de cálculo en blanco.
* Asígnale un nombre descriptivo, como "Mis Inversiones Indexa".

### 3. Abrir el Editor de Apps Script

* En tu nueva hoja de cálculo, haz clic en el menú **`Extensiones`** > **`Apps Script`**.
* Se abrirá una nueva pestaña con el editor de código. Verás un `Código.gs` con una función de ejemplo.

### 4. Pegar el Código

* Borra todo el código de ejemplo que se encuentra en `Código.gs`.
* Copia y pega el contenido del archivo `Codigo.gs` (disponible en este repositorio) en su lugar.

    **¡Importante!** Se recomienda encarecidamente revisar el código antes de usarlo. Nunca uses un fragmento de código que no hayas leído y comprendido completamente.

### 5. Guardar el Proyecto y Otorgar Permisos

* Haz clic en el **icono del disquete** ("Guardar proyecto") en la barra de herramientas del editor de Apps Script. Puedes nombrar el proyecto "Script Indexa".
* Vuelve a la pestaña de tu Google Sheet y **recarga la página** (presiona `F5` o `⌘+R`).
* Deberías ver un nuevo menú en la barra superior de tu hoja de cálculo, llamado `Indexa Capital`.
* Haz clic en **`Indexa Capital`** > **`1. Configurar Credenciales`**.
* La primera vez, Google te pedirá autorización para que el script se ejecute.
    * Haz clic en **`Revisar permisos`**.
    * Selecciona tu cuenta de Google.
    * Verás una advertencia de "Google no ha verificado esta aplicación". Esto es normal porque el código lo has creado tú. Haz clic en **`Configuración avanzada`** y luego en **`Ir a [nombre de tu script] (no seguro)`**.
    * Finalmente, haz clic en **`Permitir`**.

### 6. Configurar tus Credenciales

* Después de otorgar los permisos, el script se ejecutará automáticamente.
* Aparecerá una ventana pidiéndote el **Token de la API**. Pega el token que copiaste en el Paso 1.
* Luego, aparecerá otra ventana pidiéndote tu **ID de cuenta**. Pega el identificador que también copiaste en el Paso 1.
* Deberías recibir un mensaje de "¡Credenciales guardadas con éxito!".

### 7. ¡Generar tu Dashboard!

* Una vez configuradas las credenciales, vuelve al menú **`Indexa Capital`**.
* Haz clic en **`2. Actualizar Dashboard`**.
* Espera unos segundos. Verás una barra lateral que indica "Cargando...". Cuando termine, se habrá creado (o actualizado) una nueva hoja llamada "Dashboard" con toda tu información de inversiones, perfectamente formateada.

---

## Contribuciones

¡Las contribuciones son bienvenidas! Si encuentras alguna mejora o error, no dudes en abrir un *issue* o enviar un *pull request*.

---

## Invitación a Indexa Capital

Si aún no tienes cuenta en Indexa Capital y te gustaría abrir una, puedes usar mi enlace de invitación: <https://indexacapital.com/t/cwFqF4>
