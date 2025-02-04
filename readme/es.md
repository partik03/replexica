<p align="center">
  <a href="https://replexica.com">
    <img src="/content/banner.dark.png" width="100%" alt="Replexica" />
  </a>
</p>

<p align="center">
  <strong>⚡️ Localización con IA de última generación para web y móvil, directamente desde CI/CD.</strong>
</p>

<br />

<p align="center">
  <a href="https://replexica.com">Sitio web</a> •
  <a href="https://github.com/replexica/replexica/issues?q=is%3Aissue+is%3Aopen+label%3A%22good+first+issue%22">Contribuir</a> •
  <a href="#-github-action">GitHub Action</a> •
  <a href="#-localization-compiler-experimental">Compilador de localización</a>
</p>

<p align="center">
  <a href="https://github.com/replexica/replexica/actions/workflows/release.yml">
    <img src="https://github.com/replexica/replexica/actions/workflows/release.yml/badge.svg" alt="Release" />
  </a>
  <a href="https://github.com/replexica/replexica/blob/main/LICENSE.md">
    <img src="https://img.shields.io/github/license/replexica/replexica" alt="License" />
  </a>
  <a href="https://github.com/replexica/replexica/commits/main">
    <img src="https://img.shields.io/github/last-commit/replexica/replexica" alt="Last Commit" />
  </a>
</p>

<br />

Replexica IA automatiza la localización de software de principio a fin.

Produce traducciones auténticas al instante, eliminando el trabajo manual y la sobrecarga de gestión. El motor de localización de Replexica comprende el contexto del producto, creando traducciones perfeccionadas que los hablantes nativos esperan en más de 60 idiomas. Como resultado, los equipos realizan la localización 100 veces más rápido, con calidad de última generación, llevando funcionalidades a más clientes de pago en todo el mundo.

## 💫 Inicio rápido

1. **Solicita acceso**: [habla con nosotros](https://replexica.com/go/call) para convertirte en cliente.

2. Una vez aprobado, inicializa tu proyecto:
   ```bash
   npx replexica@latest init
   ```

3. Localiza tu contenido:
   ```bash
   npx replexica@latest i18n
   ```

## 🤖 GitHub Action

Replexica ofrece un GitHub Action para automatizar la localización en tu pipeline de CI/CD. Aquí tienes una configuración básica:

```yaml
- uses: replexica/replexica@main
  with:
    api-key: ${{ secrets.REPLEXICA_API_KEY }}
```

Esta acción ejecuta `replexica i18n` en cada push, manteniendo tus traducciones actualizadas automáticamente.

Para el modo de pull request y otras opciones de configuración, visita nuestra [documentación de GitHub Action](https://docs.replexica.com/setup/gha).

## 🧪 Compilador de localización (experimental)

Este repositorio también contiene nuestro nuevo experimento: un compilador de localización para JS/React.

Permite que los equipos de desarrollo realicen la localización frontend **sin extraer cadenas a archivos de traducción**. Los equipos pueden obtener un frontend multilingüe con solo una línea de código. Funciona en tiempo de compilación, utiliza manipulación del árbol de sintaxis abstracta (AST) y generación de código.

Puedes ver la demo [aquí](https://x.com/MaxPrilutskiy/status/1781011350136734055).

Si deseas experimentar con el compilador por tu cuenta, asegúrate de ejecutar primero `git checkout 6c6d59e8aa27836fd608f9258ea4dea82961120f`.

## 🥇 Por qué los equipos eligen Replexica

- 🔥 **Integración instantánea**: Configuración en minutos
- 🔄 **Automatización CI/CD**: Integración perfecta con el pipeline de desarrollo
- 🌍 **Más de 60 idiomas**: Expándete globalmente sin esfuerzo
- 🧠 **Motor de localización IA**: Traducciones que realmente se ajustan a tu producto
- 📊 **Flexible en formatos**: Compatible con JSON, YAML, CSV, Markdown y más

## 🛠️ Características potenciadas

- ⚡️ **Ultrarrápido**: Localización con IA en segundos
- 🔄 **Actualizaciones automáticas**: Sincronización con el contenido más reciente
- 🌟 **Calidad nativa**: Traducciones que suenan auténticas
- 👨‍💻 **Amigable para desarrolladores**: CLI que se integra con tu flujo de trabajo
- 📈 **Escalable**: Para startups en crecimiento y equipos empresariales

## 📚 Documentación

Para guías detalladas y referencias de la API, visita la [documentación](https://replexica.com/go/docs).

## 🤝 Contribuir

¿Interesado en contribuir, incluso si no eres cliente?

Revisa los [Issues para Principiantes](https://github.com/replexica/replexica/labels/good%20first%20issue) y lee la [Guía de Contribución](./CONTRIBUTING.md).

## 🧠 Equipo

- **[Veronica](https://github.com/vrcprl)**
- **[Max](https://github.com/maxprilutskiy)**

¿Preguntas o consultas? Envía un correo a veronica@replexica.com

## 🌐 Readme en otros idiomas

- [English](https://github.com/replexica/replexica)
- [Spanish](/readme/es.md)
- [French](/readme/fr.md)
- [Russian](/readme/ru.md)
- [German](/readme/de.md)
- [Chinese](/readme/zh-Hans.md)
- [Korean](/readme/ko.md)
- [Japanese](/readme/ja.md)
- [Italian](/readme/it.md)
- [Arabic](/readme/ar.md)

¿No ves tu idioma? Simplemente agrega un nuevo código de idioma al archivo [`i18n.json`](./i18n.json) y abre un PR.
