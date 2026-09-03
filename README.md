# fichaje

Fichaje de presencia por NFC. Etiqueta **NTAG216** pegada en el local, el empleado
acerca su móvil y ficha: sin app, funciona en Android y en iPhone.

- `index.html` — página de fichaje (GitHub Pages)
- `admin.html` — panel de control (horas por empleado, quién está dentro, CSV)
- `Codigo.gs` — backend de Google Apps Script (API JSONP + Google Sheets)

## Cómo se graba el tag

Con NXP TagWriter (Android), escribe una URL:

```
https://puntofibra.github.io/fichaje/?p=LOCAL01&k=SECRETO&c=000000
```

El secreto sale de la hoja **Puntos** después de ejecutar `setup()`.
En opciones avanzadas activa el **counter mirror** (contador NFC) sobre los seis
ceros y protege el tag con contraseña (solo escritura, lectura libre).

## Qué prueba el contador

La etiqueta sustituye esos seis ceros por el número de lecturas, en hexadecimal,
cada vez que alguien la acerca. El servidor exige que el contador **suba siempre**:
una URL copiada o un tag clonado se rechazan en cuanto otra persona ficha después.
No prueba *cuándo* se leyó, por eso el envío es automático al cargar y se guarda
la geolocalización.
