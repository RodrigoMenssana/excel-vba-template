# Excel VBA Template

Template profissional para desenvolvimento de projetos em **Excel VBA**, focado em organização de código, boas práticas e reutilização de estrutura.

Este repositório serve como modelo base para criação de soluções em Excel com VBA de forma escalável e organizada.

---

## 📌 Objetivo

Fornecer uma estrutura padronizada para projetos VBA que permita:

* Organização clara do código
* Versionamento eficiente com Git
* Reutilização entre projetos
* Facilidade de manutenção
* Separação entre planilha e código-fonte

---

## 🧱 Estrutura do Projeto

```
excel-vba-template/
│
├── src/
│   ├── modules/      # Módulos padrão (.bas)
│   ├── classes/      # Classes VBA (.cls)
│   └── forms/        # UserForms (.frm / .frx)
│
├── workbook/
│   └── template.xlsm # Arquivo Excel principal
│
├── docs/
│   ├── architecture.md
│   └── conventions.md
│
├── scripts/          # Scripts auxiliares
│
├── .gitignore
└── README.md
```

---

## 💡 Conceito Principal

O arquivo `.xlsm` funciona apenas como **container da aplicação**.

O código VBA é exportado para arquivos texto dentro da pasta `src/`, permitindo:

* Histórico real de alterações
* Comparação de código no Git
* Melhor controle de versão

---

## 🚀 Como Usar

### 1. Clone o repositório

```
git clone https://github.com/SEU-USUARIO/excel-vba-template.git
```

---

### 2. Abra o arquivo Excel

```
workbook/template.xlsm
```

Habilite macros ao abrir.

---

### 3. Desenvolva normalmente no VBA Editor

Abra o editor:

```
ALT + F11
```

---

### 4. Exporte os módulos VBA

No editor VBA:

```
Clique direito no módulo → Export File
```

Salve em:

```
src/modules/
src/classes/
src/forms/
```

---

## 🧭 Convenções de Nome

| Tipo   | Prefixo | Exemplo   |
| ------ | ------- | --------- |
| Module | mod     | modUtils  |
| Class  | cls     | clsLogger |
| Form   | frm     | frmLogin  |

---

## 🧩 Organização de Código Recomendada

```vba
'========================
' Public API
'========================

Public Sub RunProcess()

End Sub

'========================
' Private Helpers
'========================

Private Function ValidateInput() As Boolean

End Function
```

---

## ✅ Boas Práticas

* Evitar lógica diretamente nas planilhas
* Centralizar regras de negócio em módulos/classes
* Usar `Option Explicit`
* Separar interface e lógica
* Versionar código exportado

---

## 🔄 Workflow com Git

1. Desenvolver no Excel
2. Exportar módulos para `/src`
3. Executar:

```
git add .
git commit -m "Descrição da alteração"
git push
```

---

## 📚 Documentação

Documentos adicionais ficam em:

```
/docs
```

* Arquitetura do projeto
* Padrões de desenvolvimento
* Convenções adotadas

---

## 📄 Licença

Distribuído sob licença MIT.
Sinta-se livre para utilizar e adaptar.

---

## 👨‍💻 Autor

Rodrigo Menssana

---

## ⭐ Objetivo do Projeto

Este projeto busca aplicar conceitos de engenharia de software ao desenvolvimento em Excel VBA, transformando planilhas em aplicações estruturadas e sustentáveis.

