# Portal Gijón CBP · código SPFx

Web parts del portal interno (SharePoint Online, SPFx 1.21.1, Node 22).

```bash
cd portal
npm ci
npx gulp bundle --ship && npx gulp package-solution --ship
# paquete: sharepoint/solution/gijon-cbp-portal.sppkg
```

Antes de compilar, sube la versión en `config/package-solution.json` (dos sitios).
Despliegue: catálogo de aplicaciones del inquilino → Cargar → reemplazar `gijon-cbp-portal.sppkg` → Implementar.
Ver la guía de uso, mantenimiento y desarrollo para el detalle.
