# Gijón CBP · Sincronización FBPA

Proceso automático que cada día:

1. Lee la página del club en la [FBPA](https://www.fbpa.es/resultados-club-7954/cb-pumarin) (calendario, horarios, pistas y resultados).
2. Actualiza la lista **Calendario** del portal interno de SharePoint: crea partidos nuevos y corrige fechas, pabellones y resultados.
3. Si cambia la hora o la pista de un partido de los próximos 10 días, publica un **aviso** en el portal.
4. Publica los calendarios `.ics` de cada equipo en GitHub Pages para suscribirse desde el móvil.

## Seguridad

- El acceso a Microsoft 365 usa una **identidad federada**: no hay contraseñas ni claves guardadas en este repositorio.
- La aplicación registrada solo tiene permiso de escritura sobre el sitio del portal (`Sites.Selected`), no sobre el correo ni otros sitios.
- Los `.ics` solo contienen partidos, torneos e hitos, que ya son públicos en la FBPA. Nunca jugadores, asistencia ni información de la junta.

## Configuración

Los identificadores (no son secretos) están en `src/sync.mjs` y pueden sobrescribirse con variables del repositorio:

| Variable | Valor |
| --- | --- |
| `CLIENT_ID` | `17a79756-e595-41f7-a55e-2d616c4a9d6c` (aplicación *Gijon CBP Sync*) |
| `TENANT_ID` | `92d41dc0-249a-4073-bd9e-d966005c3a94` |

Ejecución manual: pestaña **Actions** → *Sincronizar FBPA* → *Run workflow*.

Si la FBPA cambia el diseño de su web o los nombres de los equipos, hay que ajustar `src/fbpa.mjs` (tabla `MAPA`).
