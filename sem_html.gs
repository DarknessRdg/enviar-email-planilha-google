// Adicone a URL para acessar a planilha
// Caso queira usar a planilha ativa no momento, então deixe o valor vazio.
// 
// Ex:
//    * URL_DA_PLANILHA = ""   -> utiliza a planilha ativa no momento
//    * URL_DA_PLANILHA = "https://docs.google.com/spreadsheets/...."  -> utiliza a planilha do link
const URL_DA_PLANILHA = ""

// Adicone a LETRA da coluna onde estarão todos os e-mails
const COLUNA_DO_EMAIL = "F"

// Adiciona e LETRA da coluna onde estarão todas as datas de vencimento
// OBS: a data deve estar no formato:  dia / mês / ano 
// Ex: 
//    * 1/1/2020
//    * 01/01/2020
const COLUNA_DO_VENCIMENTO = "AW"

// Adicione a mensagem que irá aparecer no assunto do e-amil
const ASSUNTO_DO_EMAIL = "Educa Damásio - Etapa em atraso"

// Adicione a mensagem que irá aparecer no corpo do e-mail
const MENSAGEM_NO_CORPO_DO_EMAIL = "Olá! Notamos que a sua etapa está em atraso. nos procure para negociarmos um novo prazo."


/** ----------- FIM DAS CONFIGURACOES ----------- */

const APP = SpreadsheetApp;
const EMAIL_APP = MailApp;


/**
 * Função principal
 */
function main() {
  const emails_atrasados = get_todos_emails_atrasados()

  for (email of emails_atrasados) {
    console.log("Enviando um e-mail para: ", email)
    envia_email(email)
  }

}


/**
 * Envia o e-mail para um e-email que está atrasado
 */
function envia_email(email) {
  EMAIL_APP.sendEmail(email, ASSUNTO_DO_EMAIL, MENSAGEM_NO_CORPO_DO_EMAIL)
}

/**
 * Retorna um Array com todos os e-mails atrasados.
 */
function get_todos_emails_atrasados() {
  console.log("---- PROCURANDO OS E-MAIL QUE ESTÃO ATRASADOS -----")
  const linhas_da_planilha = get_linhas_da_planilha()

  const emails_atrasados = []

  linhas_da_planilha.forEach( (linha, index) => {
    const email = extrair_email_da_linha(linha)
    const vencimento = extrair_vencimento_da_linha(linha)

    if (!email) {
      console.log("Linha ", index + 1, " está sem e-mail")
      return
    }

    if (esta_atrasado(vencimento)) {
        console.log("Email atrasado: ",  email, " --- Data vencimento: ", vencimento.toLocaleDateString())
        emails_atrasados.push(email)
    }
  })

  console.log("----- TODOS E-MAILS ATRASADOS ENCONTRADOS -----\n\n")
  return emails_atrasados
}


/**
 * Verifica se uma data de vencimento está atrasada.
 * Está atrase se: hoje tiver passado o dia do vencimento;
 * Se hoje for o mesmo dia do vencimento, então não está atrasdo
 */
function esta_atrasado(vencimento) {
  const hoje = new Date()

  return vencimento <= hoje
}


/**
 * Retorna uma matriz com todas as linhas e colunas da planilha
 * ex:
 *    [ 
 *       ["primeira coluna da linha 0", "segunda coluna da linha 0"],
 *       ["primeira coluna da linha 1", "segunda coluna da linha 1"],
 *       ["primeira coluna da linha 2", "segunda coluna da linha 2"],
 *       ....
 *    ]
 */
function get_linhas_da_planilha() {
  var planilha = get_planilha();

  const total_de_linhas =  planilha.getLastRow() + 1

  const emails_e_vecimentos = []

  for (let linha = 1; linha <= total_de_linhas; linha++) {
    const email = planilha.getRange(`${COLUNA_DO_EMAIL}${linha}`).getValue()
    const vencimento = planilha.getRange(`${COLUNA_DO_VENCIMENTO}${linha}`).getValue()

    emails_e_vecimentos.push([email, vencimento])
  }

  return emails_e_vecimentos
}


/**
 * Retonar a planilha do link da variável URL_DA_PLANILHA
 * 
 * caso ela seja uma string vazia: URL_DA_PLANILHA = ""
 * a planilha retorna é que estiver aberta no momento: getActiveSpreadsheet
 */
function get_planilha() {
  if (URL_DA_PLANILHA === "") {
    return APP.getActiveSpreadsheet()
  }

  return APP.openByUrl(URL_DA_PLANILHA)
}


/**
 * Retorna o maior entre dois números
 */
function max(numero_1, numero_2) {
  if (numero_1 > numero_2) {
    return numero_1
  }
  return numero_2
}


/**
 * Retorna o e-email da linha
 */
function extrair_email_da_linha(linha) {
  return linha[0]
}


/**
 * Retorna uma data do tipo Date() da linha
 */
function extrair_vencimento_da_linha(linha) {
  const vencimento = linha[1]

  if (!is_date(vencimento)) {
    return parse_data(vencimento)
  }

  return vencimento
}


/**
 * Verifica se o valor é do tipo Date()
 */
function is_date(value) {
  return value instanceof Date
}


/**
 * Transforma uma string no formato DD/MM/YYYY no tipo Date()
 */
function parse_data(data) {
  var dia  = data.split("/")[0]
  var mes  = data.split("/")[1]
  var ano  = data.split("/")[2]

  return new Date(parseInt(ano), parseInt(mes) - 1, parseInt(dia))
}
