# Mejora de acceso por área — Registro IPH

Esta versión agrega:

1. Login por área o usuario.
2. Sesión en Flask.
3. Rol `AREA` y rol `ADMIN`.
4. Filtro en `/api/registros` según el área de la sesión.
5. Validación al guardar para impedir que un área capture como otra.
6. Exportaciones filtradas automáticamente.
7. Acceso total para `ESTADISTICAS`.

## Archivos incluidos

- `app.py`: backend Flask con control de acceso.
- `Detenidos.html`: formulario con pequeño indicador de sesión y bloqueo visual de `id_operativo` para usuarios de área.
- `login.html`: pantalla de acceso.
- `usuarios_ejemplo.json`: plantilla para crear usuarios en pruebas locales.

## Uso local rápido

1. Copia `usuarios_ejemplo.json` y renómbralo como `usuarios.json`.
2. Cambia las contraseñas de ejemplo.
3. Ejecuta tu aplicación normalmente.
4. Entra al navegador; primero aparecerá la pantalla de login.

## Recomendación para Render

Lo más seguro es NO subir `usuarios.json` con contraseñas reales al repositorio.
En Render configura variables de entorno:

- `SECRET_KEY`
- `PASS_ESTADISTICAS`
- `PASS_PEP_BJ`
- `PASS_PEP_PLAYA`
- `PASS_PEP_TUL`
- `PASS_PEP_COZ`
- `PASS_PEP_OPB`
- `PASS_PEP_CAMINOS`
- `PASS_POLICIA_RURAL`
- `PASS_GPO_CENTURION`
- `PASS_GPO_ORION`
- `PASS_GPO_PDI`
- `PASS_GPO_JAGUAR`
- `PASS_MUN_FCP`
- `PASS_MUN_LC`
- `PASS_MUN_PLAYA`
- `PASS_MUN_TUL`
- `PASS_MUN_PM`
- `PASS_MUN_JMM`
- `PASS_MUN_IM`
- `PASS_MUN_COZ`
- `PASS_MUN_BJ`
- `PASS_MUN_RURAL`
- `PASS_MUN_BACALAR`

## Regla de seguridad implementada

- Si entra `PEP_TUL`, solo ve registros `PEP_TUL` y solo puede guardar como `PEP_TUL`.
- Si entra `PEP_BJ`, solo ve registros `PEP_BJ` y solo puede guardar como `PEP_BJ`.
- Si entra `ESTADISTICAS`, ve todos los registros y puede exportar todo.

## Nota importante

El archivo `registros.json` sigue siendo una solución temporal. Para trabajo real con varias oficinas y datos sensibles, después conviene migrar a una base de datos persistente.
