# Ahora Sí · Control financiero del hogar

Aplicación web para ordenar el presupuesto mensual de una familia: ingresos, gastos por categoría y ahorro por metas, todo desde el celular.

**[Ver demo con datos ficticios](https://rocioaguayo.github.io/demos/ahora-si/)** · correo `demo@ahorasi.cl` · PIN `1234`

## Qué resuelve

Llevar las cuentas de la casa en una planilla es engorroso desde el celular, y a fin de mes cuesta saber cuánto queda disponible. Ahora Sí registra cada movimiento en segundos y lo convierte en un panel para decidir:

- **Saldo disponible del mes** y cuánto se puede gastar por semana en lo que resta del mes.
- **Presupuesto por categoría**: lo planificado contra lo gastado, con aviso al 80% y al superarlo.
- **Últimos 6 meses**: ingresos, gastos y ahorro, y la variación de gastos contra el mes anterior.
- **Metas de ahorro con objetivo**: avance, monto faltante y fecha estimada según el ritmo de ahorro.
- **Movimientos editables**: corregir o eliminar registros, y descargarlos a Excel (CSV).

## Cómo funciona

| Capa | Tecnología |
|---|---|
| Interfaz | HTML, CSS y JavaScript sin librerías (un solo archivo) |
| Servidor | Google Apps Script |
| Datos | Google Sheets (hojas `Usuarios` y `Movimientos`) |
| Correo | GmailApp (bienvenida al registrarse) |

El mismo `index.html` funciona de dos formas. Dentro de Apps Script usa la planilla real. Fuera de Apps Script, por ejemplo en GitHub Pages, entra en **modo demo** con datos ficticios en memoria.

## Seguridad

- **PIN cifrado:** se guarda como hash SHA-256 con sal, nunca en texto plano. Las cuentas antiguas se migran solas la primera vez que entran.
- **Sesiones temporales:** al ingresar, el servidor entrega un token que dura 6 horas. El navegador nunca maneja el identificador interno del usuario.
- **Bloqueo por intentos:** después de 5 intentos fallidos, la cuenta queda bloqueada 15 minutos.
- **Sin datos sensibles en el código:** el ID de la planilla se guarda en las propiedades del script, fuera del repositorio.
- **Validación en el servidor:** montos, tipos y metas se revisan antes de guardar. Los textos se protegen contra inyección de fórmulas en Sheets y se muestran sin interpretar HTML.

## Instalación

1. Crea una planilla en Google Sheets con las hojas `Usuarios` y `Movimientos` (si no existen, se crean solas).
2. En la planilla: **Extensiones > Apps Script**. Copia `Code.gs`, `index.html` y `appsscript.json`.
3. En **Configuración del proyecto > Propiedades del script**, agrega `SHEET_ID` con el ID de la planilla (el texto entre `/d/` y `/edit` en la URL).
4. **Implementar > Nueva implementación > Aplicación web.**

## Estructura de datos

```
Usuarios:     id | nombre | pinHash | ingreso | metas | fechaRegistro | correo | config (JSON)
Movimientos:  id | userId | tipo | categoria | monto | descripcion | metaAhorro | fecha
```

## Próximos pasos

- Metas de ahorro creadas por el usuario.
- Movimientos recurrentes (sueldo, dividendo) que se registran solos.

---
Desarrollado por Rocío Aguayo.
