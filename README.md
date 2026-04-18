# 🧾 VBA Client & Order Management System

Sistema de gestión de clientes y pedidos desarrollado en **Excel VBA**, diseñado para automatizar procesos administrativos como el registro, actualización y consulta de información.

## 🚀 Características

- 📋 Registro de clientes
- ✏️ Actualización de datos de clientes
- 🔍 Búsqueda y filtrado de información
- 📦 Gestión de pedidos
- 🧩 Formularios interactivos (UserForms)
- 🗂️ Manejo estructurado de datos en hojas de Excel
- ⚙️ Automatización mediante macros

## 🛠️ Tecnologías utilizadas

- Microsoft Excel (.xlsm)
- VBA (Visual Basic for Applications)

## 📁 Estructura del proyecto
.
├── README.md (this file)
├── docs
├── excel
│   └── Caleido_Pedidos_Database_prefinal.xlsm
└── src
    ├── classes
    │   ├── clsDiaCalendario.cls
    │   └── clsEventosProducto.cls
    ├── forms
    │   ├── Busqueda.frm
    │   ├── Eliminar.frm
    │   ├── Entrada.frm
    │   ├── UserForm1.frm
    │   ├── frmCalendario.frm
    │   └── mostrar_imagen.frm
    └── modules
        ├── Lammado_de_Formularios.bas
        ├── ModuloCalendario.bas
        ├── ModuloClientes.bas
        ├── ModuloRegModElim.bas
        ├── ModuloValidaciones.bas
        └── ped_id_creator.bas

## 🧠 Arquitectura

El sistema sigue una estructura modular:

- **Módulos**: lógica principal del sistema
- **Formularios**: interfaz de usuario
- **Clases**: manejo de entidades (clientes, pedidos, resultados, etc.)

Se implementa separación de responsabilidades para facilitar mantenimiento y escalabilidad.

## 🎬 Demostración

*(Aquí puedes agregar GIFs grabados con OBS y convertidos)*

Ejemplo:

- Alta de cliente
- Edición de datos
- Flujo de pedidos

## ⚙️ Instalación y uso

1. Clona el repositorio:
   ```bash
   git clone https://github.com/alienfibio-25/vba_client_order_management_system_beta.git