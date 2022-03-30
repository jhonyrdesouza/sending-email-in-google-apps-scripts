# Recebendo e-mails utilizando o Google Apps Script.
> Um exemplo passo a passo de como usar um formulário HTML para enviar uma mensagem do tipo "fale conosco" por e-mail sem um servidor de back-end utilizando o serviço do Google Apps Scripts. **[Veja um exemplo funcional aqui.](jhony.me/links)**

## Por quê?
Precisava de uma maneira de receber os dados preenchidos de um formulário "Newsletters" em uma página HTML estática. Sem utilizar um servidor para realizar os envios de e-mails. Precisava dos dados para importar para a substack.com, ferramenta de boletim informativo por e-mail para pequenos escritores no qual faço envio dos meus conteúdos techs que escrevo.

## Principais vantagens:
- Sem "backend" para implantar, manterou pagar;
- Inteiramente customizável - cada aspecto é customizável!
- E- mail enviado via Google Mail que está na lista de permissões em todos os lugares (sucesso de alta capacidade de entrega);
- Colete/armazene quaisquer dados de formulário em uma planilha para facilitar a visualização (perfeito se você precisar compartilhá-lo com pessoas não técnicas);

## Script:

``` js
// E-mail ao qual será enviado as respostas do formulário;
var TO_ADDRESS = "seu@email.com";

// Copia todas as chaves/valores do formulário em HTML para o e-mail;
// Usa um array de chaves se fornecido, ou, o objeto para determinar a ordem dos campos;
function formatMailBody(obj, order) {
  var result = "";
  if (!order) {
    order = Object.keys(obj);
  }

  // Faz um loop sobre todas as chaves nos dados do formulário ordenado;
  for (var idx in order) {
    var key = order[idx];
    result +=
      "<h4 style='text-transform: capitalize; margin-bottom: 0'>" +
      key +
      "</h4><div>" +
      sanitizeInput(obj[key]) +
      "</div>"; // Para cada chave, concatenar um par `<h4/>`/`<div/>` do nome da chave e seu valor e anexá-lo à string `result` criada no início;
  }
  return result;
  // Uma vez que o loop é feito, `result` será uma longa string para colocar no corpo do email;
}

// Limpa o conteúdo do usuário;
function sanitizeInput(rawInput) {
  var placeholder = HtmlService.createHtmlOutput(" ");
  placeholder.appendUntrusted(rawInput);
  return placeholder.getContent();
}

function doPost(e) {
  try {
    Logger.log(e); // A versão do Google Script de console.log consulte: Class Logger;
    record_data(e);

    // Nome mais curto para os dados de formulário;
    var mailData = e.parameters;

    // Nomes e ordem dos elementos do formulário (se definido);
    var orderParameter = e.parameters.formDataNameOrder;
    var dataOrder;
    if (orderParameter) {
      dataOrder = JSON.parse(orderParameter);
    }

    // Determina o destinatário do e-mail;
    // Se você tiver seu e-mail sem comentários acima, ele usa esse `TO_ADDRESS`
    // Caso contrário, o padrão é o email fornecido pelo atributo de dados do formulário;
    var sendEmailTo =
      typeof TO_ADDRESS !== "undefined"
        ? TO_ADDRESS
        : mailData.formGoogleSendEmail;

    // Envia e-mail se o endereço estiver definido
    if (sendEmailTo) {
      MailApp.sendEmail({
        to: String(sendEmailTo),
        subject: "Notificação do forumulário da Newsletters | jhony.me/links",
        replyTo: String(mailData.email), // Isso é opcional e depende do seu formulário realmente coletando um campo chamado `email`;
        htmlBody: formatMailBody(mailData, dataOrder)
      });
    }

    return ContentService.createTextOutput( // Retorna resultados de sucesso do JSON;
      JSON.stringify({ result: "success", data: JSON.stringify(e.parameters) })
    )
      .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    // Se erro retornar isso;
    Logger.log(error);
    return ContentService.createTextOutput(
      JSON.stringify({ result: "error", error: error })
    ).setMimeType(ContentService.MimeType.JSON);
  }
}

// Record_data insere os dados recebidos do envio do formulário html e são os dados recebidos do POST;
function record_data(e) {
  var lock = LockService.getDocumentLock();
  lock.waitLock(30000); // Espera até 30 segundos para evitar escrita simultânea;
  try {
    Logger.log(JSON.stringify(e)); // Registra os dados POST caso precisemos depurá-los;

    // Seleciona a planilha 'respostas' por padrão;
    var doc = SpreadsheetApp.getActiveSpreadsheet();
    var sheetName = e.parameters.formGoogleSheetName || "responses";
    var sheet = doc.getSheetByName(sheetName);

    var oldHeader = sheet
      .getRange(1, 1, 1, sheet.getLastColumn())
      .getValues()[0];
    var newHeader = oldHeader.slice();
    var fieldsFromForm = getDataColumns(e.parameters);
    var row = [new Date()]; // O primeiro elemento da linha deve ser sempre um timestamp sendo dd/mm/aaaa hh:mm:ss;
    

    // Percorre as colunas do cabeçalho;
    for (var i = 1; i < oldHeader.length; i++) {
      
      // Começa em 1 para evitar a coluna Timestamp;
      var field = oldHeader[i];
      var output = getFieldFromData(field, e.parameters);
      row.push(output);

      // Marca como armazenado removendo dos campos do formulário;
      var formIndex = fieldsFromForm.indexOf(field);
      if (formIndex > -1) {
        fieldsFromForm.splice(formIndex, 1);
      }
    }

    // Define quaisquer novos campos em nosso formulário;
    for (var i = 0; i < fieldsFromForm.length; i++) {
      var field = fieldsFromForm[i];
      var output = getFieldFromData(field, e.parameters);
      row.push(output);
      newHeader.push(field);
    }

    // Mais eficiente para definir valores como array [][] do que individualmente;
    var nextRow = sheet.getLastRow() + 1; // Obtém a próxima linha;
    sheet.getRange(nextRow, 1, 1, row.length).setValues([row]);

    // Atualiza a linha do cabeçalho com quaisquer novos dados;
    if (newHeader.length > oldHeader.length) {
      sheet.getRange(1, 1, 1, newHeader.length).setValues([newHeader]);
    }
  } catch (error) {
    Logger.log(error);
  } finally {
    lock.releaseLock();
    return;
  }
}

function getDataColumns(data) {
  return Object.keys(data).filter(function (column) {
    return !(
      column === "formDataNameOrder" ||
      column === "formGoogleSheetName" ||
      column === "formGoogleSendEmail" ||
      column === "honeypot"
    );
  });
}

function getFieldFromData(field, data) {
  var values = data[field] || "";
  var output = values.join ? values.join(", ") : values;
  return output;
}
```



### Observação:
Com a LGPD, recomendo pesquisar recomendações sobre privacidade do usuário; você pode ser responsabilizado pela guarda dos dados pessoais dos usuários e deve fornecer a eles uma maneira de entrar em contato com você. :)

### Aviso:
A API do Google tem limites de quantos e-mails pode enviar por dia. Isso pode variar na sua conta do Google, [veja os limites aqui.](https://developers.google.com/apps-script/guides/services/quotas). Recomendo implementar este tutorial até a Parte 3, pois os dados sempre serão adicionados à planilha primeiro e depois enviados por e-mail, se possível.

### Licença:

The [MIT License]() (MIT)

Copyright :copyright: 2022 - Recebendo e-mails utilizando o Google Apps Script.


