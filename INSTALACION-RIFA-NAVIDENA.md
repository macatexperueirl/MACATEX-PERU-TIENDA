# Instalación de la Rifa Navideña MACATEXPERU 2026

## Estado
El módulo de interfaz está integrado en `index.html` y usa `rifa-navidad.js`. El backend está preparado en `RifaNavidad.gs`, pero debe incorporarse al proyecto de Google Apps Script que actualmente atiende el catálogo y publicarse como una nueva versión antes de que las reservas funcionen en producción.

## 1. Incorporar el backend
1. Abre el proyecto de Google Apps Script asociado a la hoja de cálculo que contiene PRODUCTOS, CANASTAS y DETALLE_CANASTAS.
2. Crea un archivo de script llamado `RifaNavidad.gs` y copia el contenido de este repositorio.
3. En la función existente `doGet(e)`, justo después de obtener la variable `tipo`, agrega esta condición:

```javascript
if (tipo === 'rifa') return salida(obtenerRifaPublica_(), e);
```

4. Guarda el proyecto. El archivo incluye `doPost(e)`; el código compartido anteriormente no tenía una función doPost.

## 2. Crear la hoja de control
1. En el editor de Apps Script, selecciona y ejecuta `prepararRifaNavidena2026`.
2. Autoriza los permisos solicitados.
3. Verifica que se haya creado la pestaña `RIFA_2026` con 300 filas, del 001 al 300.
4. No borres ni renombres los encabezados de esa hoja.

## 3. Publicar la nueva versión
1. En Apps Script, abre **Implementar > Gestionar implementaciones**.
2. Edita la implementación de aplicación web que usa la página y selecciona **Nueva versión**.
3. Mantén el acceso configurado para que los visitantes de la página puedan consultar y solicitar reservas, según la política de acceso que hayas decidido para el proyecto.
4. Guarda la implementación. La URL debe seguir siendo la misma que está configurada en `rifa-navidad.js`.

## 4. Administración de pagos
- Para confirmar un Yape ya verificado, ejecuta `confirmarNumeroRifa` desde Apps Script y escribe el número reservado.
- Para liberar manualmente un número reservado, ejecuta `liberarNumeroRifa`.
- Las reservas vencidas se liberan cuando se consulta la disponibilidad pública o al ejecutar `actualizarRifaNavidena2026`.
- Nunca marques un número como vendido solo por recibir una captura: verifica el abono en la cuenta Yape.

## 5. Pruebas antes de anunciar
- Reserva un número de prueba y comprueba que cambie a RESERVADO.
- Comprueba que otro visitante ya no pueda reservarlo.
- Confirma el pago de prueba desde Apps Script y verifica que aparezca como VENDIDO.
- Prueba una reserva vencida y confirma que vuelva a DISPONIBLE.
- Verifica que el listado público solo muestre número y estado, sin nombre ni WhatsApp.

## Parámetros definidos
- 300 números: 001–300.
- S/10 por número.
- Yape: 930 523 654.
- Reserva: 30 minutos.
- Premios: CAN-015 Selección Especial (valor referencial S/450), juego de ollas Record y mini licuadora portátil Oster.
- Sorteo anunciado en el arte: domingo 27 de diciembre de 2026, 8:00 p. m.

**Importante:** no iniciar la venta ni promocionar la rifa como habilitada hasta completar la integración, probar el flujo y validar las bases y autorizaciones que correspondan.
