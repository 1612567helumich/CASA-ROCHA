# Casa Rocha

| Archivo | Contenido |
|---|---|
| `001_PLANO_CROCHA_R1.6.dxf` | Plano arquitectónico corregido (plantas, cortes, escalera y layouts A-01 a A-03). |
| `001_CASA_ROCHA_R1.0.e2k` | Modelo estructural para ETABS. Abrir con *File › Import › ETABS .e2k*. |
| `001_CASA_ROCHA_RESUMEN_R1.0.pdf` | Resumen breve de cambios, criterios y pendientes. |
| `originales/` | Plano R1.5 de partida. |
| `scripts/` | Generadores en Python que rehacen el DXF, el `.e2k` y el PDF. |

## Regenerar

```bash
pip install ezdxf shapely matplotlib reportlab
python3 scripts/build_dxf.py                                  # plano R1.6
python3 scripts/build_e2k.py <ruta>/002_HUAYACAN_R2.2.e2k     # modelo (usa el machote)
python3 scripts/valida_e2k.py 001_CASA_ROCHA_R1.0.e2k         # revisión de referencias del .e2k
```

Los datos supuestos del sismo (a0r, terreno, Q, R) están al inicio de `scripts/build_e2k.py`.
Las cargas también están ahí.
