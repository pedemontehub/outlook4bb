# Brew Brother para Outlook

Los archivos del complemento de Outlook de **Brew Brother Control**, publicados con GitHub Pages en
<https://pedemontehub.github.io/outlook4bb/>: el botón **New case** de un correo abierto, que abre Brew Brother con un
caso nuevo escrito desde el correo.

**No se editan aquí.** Salen de la carpeta `outlook/` del repositorio de la app (`pedemontehub/brewbro-buper`), con
`python outlook/armar.py`; para cambiarlos se vuelve a armar allí y se reemplazan estos.

- `manifest.xml`: el que se instala en Outlook (https://aka.ms/olksideload > My add-ins > Add a custom add-in > Add
  from file, o desde Microsoft 365 admin center > Integrated apps para todo el equipo).
- `commands.html`, `commands.js`: lo que corre al presionar el botón. Lee el asunto, el remitente, la fecha y el texto
  del correo y abre New case con eso en la dirección: el correo no pasa por ningún servidor.
- `icon-*.png`: el ojo de la marca.

No hay secretos: solo la dirección de la app, que pide iniciar sesión igual.
