# 🍱 Organizador e Avaliador de Marmitas & Produtos

> Aplicativo web leve, responsivo e sem servidor tradicional (serverless), criado para catalogar, classificar e avaliar refeições congeladas ou produtos de forma rápida, com sincronização em nuvem via **Google Planilhas (Google Sheets)** e funcionamento offline no navegador.

Desenvolvido por **Ana Laura A. Queiroz** via **Gemini Canvas**.

---

## 🌟 Principais Recursos

- **⚡ Autônomo e de Arquivo Único (Single-File):** Funciona abrindo diretamente o arquivo HTML no navegador, sem precisar instalar Node.js, compilar código ou configurar servidores.
- **☁️ Integração com Google Sheets (Apps Script):** 
  - Salva em tempo real em segundo plano.
  - Botão de envio manual (**📤 Enviar para Planilha**) com confirmação protetiva.
  - Botão de restauração (**📥 Puxar da Planilha**) para carregar o histórico de qualquer dispositivo.
- **⭐ Sistema Completo de Avaliação:**
  - Notas de 1 a 5 estrelas.
  - Classificação de momento ideal: **☀️ Almoço**, **🌙 Jantar** ou **✨ Tanto faz**.
  - Decisão rápida de recompra: **👍 Sim**, **👎 Não** ou **⏳ Não provada**.
  - Campo de observações livres para texturas, tempero e percepções.
- **🔍 Filtros e Busca Inteligente:**
  - Filtro por sabor, ingrediente ou anotação.
  - Abas rápidas (*Só Almoço*, *Só Jantar*, *Top 4-5 ⭐*, *Não Repetiria*, *Não Provados*).
  - Ordenação flexível (*Ordem do Catálogo*, *Maior Nota*, *Preço*, *Estoque*).
- **💾 Segurança de Dados & Portabilidade:**
  - Exportação em **CSV** para planilhas.
  - Cópia formatada para a área de transferência (**Ctrl+V** direto no Excel/Google Sheets).
  - **Backup & Restauração em JSON** com 1 clique.
  - Tolerante a atualizações de página (`F5`) mesmo quando incorporado dentro de `iframes` (ex: Google Sites, Notion ou WordPress).

---

## 📁 Estrutura do Repositório

```text
├── index.html                     # Versão pronta para uso com catálogo pré-carregado
├── template_compartilhavel.html   # Template neutro para qualquer pessoa personalizar
└── README.md                      # Documentação completa do projeto
```

---

## 🚀 Como Usar o Template no seu Próprio Google Drive

Para usar o app conectado à sua própria conta Google:

### 1. Criar a Planilha Google
1. Acesse [sheets.new](https://sheets.new) para abrir uma nova planilha em branco no seu Google Drive.
2. No menu superior da planilha, clique em **Extensões** > **Apps Script**.

### 2. Configurar o Apps Script
1. Apague qualquer código presente no editor do Apps Script.
2. Copie e cole o script abaixo:

```javascript
function doGet(e) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var lastRow = sheet.getLastRow();
    var callback = (e && e.parameter && e.parameter.callback) ? e.parameter.callback : null;
    
    if (lastRow <= 1) {
      var emptyRes = JSON.stringify({ status: "success", meals: [] });
      return callback ? ContentService.createTextOutput(callback + "(" + emptyRes + ")").setMimeType(ContentService.MimeType.JAVASCRIPT) : ContentService.createTextOutput(emptyRes).setMimeType(ContentService.MimeType.JSON);
    }
    
    var data = sheet.getRange(2, 1, lastRow - 1, 10).getValues();
    var meals = [];
    
    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      if (!row[0] || String(row[0]).trim() === "") continue;
      
      var mealTimeVal = null;
      var strTime = String(row[5] || "");
      if (strTime.indexOf("Almoço") > -1) mealTimeVal = "lunch";
      else if (strTime.indexOf("Jantar") > -1) mealTimeVal = "dinner";
      else if (strTime.indexOf("Tanto faz") > -1) mealTimeVal = "anytime";

      var repeatVal = null;
      var strRepeat = String(row[7] || "");
      if (strRepeat === "Sim") repeatVal = true;
      else if (strRepeat === "Não") repeatVal = false;

      var ratingVal = parseInt(String(row[6] || "").replace(/[^0-9]/g, "")) || 0;
      var rawPrice = String(row[4] || "").replace("R$", "").replace(".", "").replace(",", ".").trim();
      var priceVal = parseFloat(rawPrice) || 0;

      meals.push({
        id: "m_sheet_" + i + "_" + new Date().getTime(),
        name: String(row[0] || ""),
        category: String(row[1] || ""),
        quantity: parseInt(row[2]) || 1,
        weight: String(row[3] || "300g"),
        priceSingle: priceVal,
        mealTime: mealTimeVal,
        rating: ratingVal,
        repeat: repeatVal,
        notes: String(row[8] || ""),
        tag: String(row[1] || "Item")
      });
    }
    
    var outputJson = JSON.stringify({ status: "success", meals: meals });
    return callback ? ContentService.createTextOutput(callback + "(" + outputJson + ")").setMimeType(ContentService.MimeType.JAVASCRIPT) : ContentService.createTextOutput(outputJson).setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    var errJson = JSON.stringify({ status: "error", message: err.toString() });
    return callback ? ContentService.createTextOutput(callback + "(" + errJson + ")").setMimeType(ContentService.MimeType.JAVASCRIPT) : ContentService.createTextOutput(errJson).setMimeType(ContentService.MimeType.JSON);
  }
}

function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return ContentService.createTextOutput(JSON.stringify({ status: "error", message: "Sem dados" })).setMimeType(ContentService.MimeType.JSON);
    }
    var data = JSON.parse(e.postData.contents);
    
    if (!data.meals || !Array.isArray(data.meals) || data.meals.length === 0) {
      return ContentService.createTextOutput(JSON.stringify({ status: "ignored_empty" })).setMimeType(ContentService.MimeType.JSON);
    }

    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(["Prato / Sabor", "Categoria", "Qtd", "Gramatura", "Preço Avulso", "Momento Ideal", "Nota ⭐", "Repetiria?", "Anotações", "Última Atualização"]);
      var headerRange = sheet.getRange(1, 1, 1, 10);
      headerRange.setBackground("#059669").setFontColor("#FFFFFF").setFontWeight("bold");
      sheet.setFrozenRows(1);
    } else {
      var lastRow = sheet.getLastRow();
      if (lastRow > 1) {
        sheet.getRange(2, 1, lastRow - 1, 10).clearContent();
      }
    }
    
    var rows = [];
    var now = new Date().toLocaleString("pt-BR", { timeZone: "America/Sao_Paulo" });
    
    data.meals.forEach(function(m) {
      rows.push([
        m.name,
        m.category || "",
        m.quantity || 1,
        m.weight || "300g",
        "R$ " + Number(m.priceSingle || 0).toFixed(2).replace(".", ","),
        m.mealTime || "Não definido",
        (m.rating && m.rating > 0) ? m.rating + " ⭐" : "Sem nota",
        m.repeat || "Ainda não provado",
        m.notes || "",
        now
      ]);
    });
    
    if (rows.length > 0) {
      sheet.getRange(2, 1, rows.length, 10).setValues(rows);
      sheet.autoResizeColumns(1, 10);
    }
    
    return ContentService.createTextOutput(JSON.stringify({ status: "success" })).setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: err.toString() })).setMimeType(ContentService.MimeType.JSON);
  }
}
```

### 3. Publicar o Web App
1. Clique no botão azul **Implantar** (canto superior direito) > **Nova implantação**.
2. No ícone de engrenagem ⚙️, escolha **Aplicativo da Web**.
3. Em *Executar como*, mantenha **Eu**.
4. Em *Quem tem acesso*, selecione **Qualquer pessoa** (necessário para que o navegador consiga salvar os dados).
5. Clique em **Implantar**, conceda as permissões da sua conta e copie a **URL do aplicativo da Web** (que termina em `/exec`).

### 4. Conectar no Aplicativo
Abra o arquivo `template_compartilhavel.html` no seu navegador, clique no botão **⚙️ Configurar Planilha** no cabeçalho e cole a URL gerada.

---

## 💻 Como Embutir no Google Sites ou Páginas Externas

Se você hospedar o aplicativo em plataformas que usam `iframes` com isolamento restrito (como o **Google Sites**), basta abrir o arquivo `.html` em qualquer editor de texto e preencher as constantes no início da tag `<script>`:

```javascript
const HARDCODED_WEBHOOK_URL = "https://script.google.com/macros/s/SEU_SCRIPT_ID/exec"; 
const HARDCODED_SPREADSHEET_URL = "https://docs.google.com/spreadsheets/d/SEU_SHEET_ID/edit"; 
```

Dessa forma, o aplicativo já iniciará permanentemente conectado, imune à limpeza de cache ou atualizações de página (`F5`).

---

## 🛠️ Tecnologias Utilizadas

- **HTML5 & Vanilla JavaScript (ES6+):** Arquitetura sem dependências pesadas, ultra rápida e portável.
- **Tailwind CSS (CDN):** Interface moderna e responsiva em dispositivos móveis e desktops.
- **Font Awesome 6:** Iconografia de status, avaliações e utilidades.
- **Google Apps Script & Sheets API:** Armazenamento em nuvem gratuito e personalizável via endpoints Web App.

---

## 📄 Licença e Créditos

Distribuído sob a licença MIT. Consulte `LICENSE` para mais detalhes.

Projeto concebido e desenvolvido por **Ana Laura A. Queiroz** via **Gemini Canvas**.