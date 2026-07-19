# Factures JSON

Web app estàtica, pensada per a mòbil, per registrar factures emeses i rebudes, conservar-les localment i exportar-les a JSON per importar-les després a Excel.

## Funcions

- Formulari diferenciat per a factures `Emesa` i `Rebuda`.
- Persistència local amb `localStorage`.
- Edició, eliminació i control bàsic de duplicats.
- Exportació i reimportació de còpies JSON.
- OCR opcional al navegador amb Tesseract.js (`spa+cat`).
- PWA instal·lable i aplicació estàtica compatible amb GitHub Pages.
- Sense servidor ni base de dades. Les imatges no s’envien a cap API.

## Publicar amb GitHub Pages

1. Ves a **Settings → Pages**.
2. A **Build and deployment**, tria **Deploy from a branch**.
3. Selecciona la branca `main` i la carpeta `/ (root)`.
4. Desa els canvis.

## JSON exportat

```json
{
  "schema_version": 1,
  "export_id": "uuid",
  "created_at": "2026-07-19T20:30:00.000Z",
  "factures": [
    {
      "id_extern": "uuid",
      "tipus": "Rebuda",
      "data": "2026-07-19",
      "num_factura": "F-4271",
      "descripcio": "Material de conservació",
      "nom_alias": "Proveïdor",
      "nif": "B12345678",
      "bi": 125.5,
      "iva_pct": 0.21,
      "irpf_pct": null,
      "afectacio_pct": 1,
      "observacions": ""
    }
  ]
}
```

## OCR

L’OCR és una ajuda, no una importació cega. Busca data, número de factura, NIF/CIF, base imposable, tipus d’IVA i retenció. Les factures tenen formats massa variats perquè una expressió regular es converteixi sobtadament en assessoria fiscal, de manera que tots els camps s’han de revisar abans de guardar.

La primera càrrega de l’OCR necessita connexió per descarregar el motor i els idiomes. Després el navegador normalment els conserva a la memòria cau.
