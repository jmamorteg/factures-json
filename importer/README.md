# Importador JSON → Excel

Importa els JSON generats per **Factures JSON** directament a `tblEntradaFactures`, sense OneDrive, Power Automate ni convertir el llibre a `.xlsm`.

## Requisits

- Windows.
- Excel d’escriptori instal·lat.
- El llibre de control fiscal amb:
  - full `Entrada_Factures`;
  - taula `tblEntradaFactures`;
  - columnes `Estat_Fiscal`, `Tipus`, `Data`, `Num_Factura`, `Nom_Alias`, `%IVA`, `BI`, `%IRPF/%Reper`, `Observacions`, `Origen` i `ID_Extern`.

## Ús normal

1. Tanca el llibre d’Excel.
2. Fes doble clic a `Importar_factures.bat`.
3. Selecciona el JSON exportat des del mòbil.
4. Selecciona el llibre `.xlsx`.
5. Obre el llibre i revisa les files noves amb `Estat_Fiscal = Pendent`.

L’importador:

- crea una còpia de seguretat al costat del llibre;
- afegeix les files a `tblEntradaFactures`;
- conserva i estén les fórmules, formats i validacions de la taula;
- escriu `Origen = JSON`;
- conserva `ID_Extern`;
- evita duplicats pel mateix `ID_Extern`;
- evita també coincidències de `Tipus + Nom_Alias + Num_Factura + Data + BI`;
- crea un fitxer `.log` al costat del JSON.

## Mode de prova

No modifica el llibre:

```bat
Importar_factures.bat -DryRun
```

## Ús amb rutes explícites

```bat
Importar_factures.bat -JsonPath "C:\Users\nom\Downloads\factures_2026-07-19_2030.json" -WorkbookPath "C:\Comptabilitat\ControlFiscal.xlsx"
```

## Correspondència de camps

| JSON | Excel |
|---|---|
| `tipus` | `Tipus` |
| `data` | `Data` |
| `num_factura` | `Num_Factura` |
| `nom_alias` | `Nom_Alias` |
| `nif` | `NIF_Entrada`, si existeix; si no, s’afegeix a `Observacions` |
| `bi` | `BI` |
| `iva_pct` | `%IVA` |
| `irpf_pct` si és emesa | `%IRPF/%Reper` |
| `afectacio_pct` si és rebuda | `%IRPF/%Reper` |
| `descripcio` + `observacions` | `Observacions` |
| `id_extern` | `ID_Extern` |

Les columnes `Exercici_Fiscal`, `Trimestre`, imports calculats, classificació fiscal, revisió i `Computable_Fiscal` no s’escriuen: les calcula Excel.

## Nota sobre el NIF

`NIF_CIF` és una columna calculada en el llibre actual: Excel l’obté des de `Proveidors` a partir de `Nom_Alias`. Per això l’importador no la sobreescriu. Si vols conservar el NIF introduït al mòbil en una columna pròpia, afegeix una columna manual anomenada `NIF_Entrada`; en cas contrari quedarà anotat a `Observacions` per poder-lo revisar.
