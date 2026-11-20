# Repostería Sucre — sistema local para dos empleados

Aplicación en español para una tienda de insumos de repostería en Sucre. Importes en Bs y horario de Bolivia. Funciona sin internet después de instalar las dependencias. Recursos visuales locales.

## Alcance de la entrega por Bs 1.000

- Ventas con búsqueda de productos, stock automático y cálculo de cambio.
- Efectivo, QR, transferencia y tarjeta: registro manual del pago y su referencia. El operador verifica el cobro.
- Inventario por presentación (ejemplo: bolsa de 1 kg), reposición, precios y alertas de vencimiento y stock mínimo.
- Envíos con destinatario, teléfono, dirección, referencia y tarifa fija configurable. Estados pendiente, en camino y entregado.
- Dos cuentas: ventas administra productos, reportes y configuración; envíos consulta y actualiza las entregas.
- Anulaciones con devolución de stock, nota de venta imprimible y reportes exportables a Excel.
- Respaldo descargable desde Configuración y cambio de contraseña desde el nombre del usuario.

Las notas son comprobantes internos. La versión no integra facturación electrónica, confirmación bancaria automática, contabilidad, rutas GPS ni ventas fraccionadas por peso. Las ventas con envío se registran pagadas; no incluye cobro contra entrega. El vencimiento se registra por producto, sin gestión de lotes.

## Instalación en dos computadoras Windows

1. Elija una computadora principal, que permanecerá encendida durante el trabajo. Copie allí la carpeta del sistema. Instale Python 3.11 o superior y ejecute `instalar_windows.bat` con internet. La segunda computadora solo necesita navegador.
2. Ejecute `iniciar_windows.bat` en la principal. Abra `http://localhost:5000`.
3. Conecte ambas computadoras a la misma red del negocio. En la principal ejecute `ipconfig` y anote la dirección IPv4 del adaptador conectado (por ejemplo `192.168.1.20`).
4. Permita Python o el puerto TCP 5000 en el firewall únicamente para la red privada del negocio. No abra puertos del router hacia internet.
5. En la segunda computadora abra `http://192.168.1.20:5000`, sustituyendo la IP por la real. Cree un acceso directo del navegador en ambos escritorios. Reserve esa IP en el router para evitar cambios de dirección.
6. Ingrese con las cuentas de abajo y cambie ambas contraseñas. Configure nombre, dirección, teléfono y tarifa del negocio.

No ejecute dos servidores ni copie bases entre equipos para trabajar: ambas computadoras usan la única base de la principal. Si esta se apaga o se pierde la red, el segundo equipo no podrá trabajar hasta restablecer la conexión. La conexión HTTP está prevista para una red privada de confianza.

## Acceso inicial

| Usuario | Contraseña inicial | Función |
|---|---|---|
| ventas | ventas123 | Ventas y administración |
| envios | envios123 | Gestión de entregas |

La base se crea al primer inicio con productos de ejemplo; antes de entregar al cliente, ajuste o desactive estos productos y elimine las ventas de prueba desde Configuración. No cargue ventas ficticias durante la operación real.

## Linux

Ejecute `./iniciar.sh` con Python 3.11 o superior. La primera instalación requiere internet. El servidor Waitress escucha en el puerto 5000. Puede configurar `HOST`, `PORT` y `REPOSTERIA_DB` mediante variables de entorno.

## Sitio compartido en Vercel (varios negocios)

Con la variable `DATABASE_URL`, el sistema usa Postgres (Neon) en lugar de SQLite. Así funciona en Vercel. Cada negocio tiene sus propios productos, ventas, envíos, configuración y respaldo. Un usuario solo ve los datos de su negocio. En la nube no se crean las cuentas locales `ventas` ni `envios`.

1. Inicie sesión: `npx vercel login`.
2. Cree el proyecto: `npx vercel link`.
3. En el panel de Vercel abra el proyecto, vaya a Storage → Create Database → Neon y conéctela al proyecto. Elija la región São Paulo (sa-east-1), la misma que la función (`gru1` en `vercel.json`). Esto crea `DATABASE_URL`.
4. Cree la clave de sesiones: `openssl rand -hex 32 | npx vercel env add SECRET_KEY production`.
5. Publique: `npx vercel --prod`.
6. Cree en Vercel los secretos `CARITO_PASSWORD` y `JUANITA_PASSWORD`, con al menos ocho caracteres, para Producción y Preview. En el siguiente despliegue se crean una sola vez los usuarios `reposteria_carito` y `reposteria_juanita`. Las contraseñas nunca se guardan en Git. Si después se cambian desde la aplicación, los despliegues posteriores no las restablecen.

Para dar de alta otros negocios, descargue temporalmente las variables y use `crear_negocio.py`. La contraseña se muestra una sola vez:
`npx vercel env pull .env.produccion --environment=production`
`set -a; . ./.env.produccion; set +a; .venv/bin/python crear_negocio.py "Nombre del negocio" usuario`
`rm .env.produccion`

`.vercelignore` impide subir las bases `.db`, `.secret_key` y `.venv`. En la nube, el respaldo descarga un archivo `.db` con los datos de ese negocio solamente. Los números de venta son correlativos para todo el sitio, por eso cada negocio verá saltos en la numeración.

Para probar contra Postgres: `TEST_DATABASE_URL=postgresql://... .venv/bin/python -m unittest discover -s tests`. Use una base de prueba, porque se borra por completo.

## Respaldo y recuperación

Al finalizar cada jornada, desde la cuenta ventas abra Configuración → Descargar respaldo. Guarde una copia en USB y mantenga al menos las últimas siete jornadas. La descarga usa la API de backup de SQLite y se puede realizar con el sistema encendido.

Para restaurar: cierre el servidor, conserve una copia de la base actual y reemplace `reposteria.db` por el respaldo elegido. Reinicie y compruebe las últimas ventas. También guarde `.secret_key` si necesita conservar las sesiones; sin ella se genera otra clave y se vuelve a iniciar sesión. Los respaldos contienen datos de clientes y contraseñas cifradas mediante hash; guárdelos con acceso restringido.

## Comprobación antes de entregar

Ejecute `.venv/bin/python -m unittest discover -s tests -v` en Linux o `.venv\Scripts\python.exe -m unittest discover -s tests -v` en Windows. Estas pruebas usan una base temporal sin modificar la del negocio.

Verifique presencialmente desde las dos computadoras: conexión, permisos, venta con ticket, envío visible en el segundo equipo, cierre de caja y descarga y restauración de un respaldo. Pruebe la impresora real con papel de 80 mm o ajuste el navegador al formato disponible. La instalación física, firewall e impresión quedan pendientes hasta disponer de los equipos del cliente.
