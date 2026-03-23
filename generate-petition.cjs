const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, BorderStyle, WidthType, ShadingType,
  PageNumber, LevelFormat
} = require("docx");
// ╔══════════════════════════════════════════════════════════════╗
// ║  DADOS DO CASO — PREENCHA AQUI OS DADOS DO(A) AGENTE       ║
// ║  So mude o texto entre aspas "..."                          ║
// ║  Nao apague aspas, virgulas ou chaves!                      ║
// ╚══════════════════════════════════════════════════════════════╝
const DADOS = {
  // ----- AUTOR(A) — Agente de Saude -----
  autorNome: "MARIA JOSE DA SILVA",
  autorNacionalidade: "brasileira",
  autorEstadoCivil: "casada",
  autorProfissao: "Agente Comunitaria de Saude",   // "Agente Comunitario(a) de Saude" OU "Agente de Combate as Endemias"
  autorRG: "MG-12.345.678",
  autorCPF: "123.456.789-00",
  autorEndereco: "Rua das Flores, n. 123",
  autorBairro: "Centro",
  autorCidade: "Ponte Nova",
  autorCEP: "35.430-000",
  autorTelefone: "(31) 99999-8888",
  autorEmail: "maria.silva@email.com",
  // ----- TIPO DE AGENTE -----
  tipoAgente: "ACS",          // "ACS" ou "ACE"
  tipoAgentePorExtenso: "Agente Comunitaria de Saude",  // por extenso
  // ----- MUNICIPIO REU -----
  municipioNome: "PONTE NOVA",
  municipioEndereco: "Avenida Caetano Marinho, n. 306",
  municipioBairro: "Centro",
  municipioCEP: "35.430-001",
  municipioTelefone: "(31) 3819-5454",
  municipioCNPJ: "23.804.149/0001-29",
  // ----- COMARCA -----
  comarca: "PONTE NOVA",
  // ----- DADOS FUNCIONAIS -----
  dataAdmissao: "15/03/2018",
  matricula: "12345",
  lotacao: "UBS Sao Jose - PSF Centro",
  salarioBase: "2.824,00",
  // ----- PERIODO SEM REPASSE -----
  periodoInicio: "janeiro/2020",
  periodoFim: "dezembro/2024",
  mesesDevidos: "60",
  // ----- VALORES -----
  valorMensal: "350,00",        // valor mensal estimado do IFA
  valorTotal: "21.000,00",      // valor mensal x meses devidos
  valorCausa: "21.000,00",      // mesmo valor (deve ser ate 60 salarios minimos)
};
// ========================================
// HELPERS
// ========================================
function txt(text, opts = {}) {
  return new TextRun({ text, font: "Times New Roman", size: 24, ...opts });
}
function bold(text, opts = {}) {
  return txt(text, { bold: true, ...opts });
}
function emptyLine() {
  return new Paragraph({ spacing: { before: 0, after: 0 }, children: [txt("")] });
}
function heading(text) {
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 280, after: 120 },
    children: [bold(text, { size: 24 })],
  });
}
function subHeading(text) {
  return new Paragraph({
    spacing: { before: 200, after: 80 },
    indent: { firstLine: 0 },
    children: [bold(text, { size: 22 })],
  });
}
function para(runs, opts = {}) {
  return new Paragraph({
    alignment: AlignmentType.JUSTIFIED,
    spacing: { after: 120, line: 360 },
    indent: { firstLine: 720 },
    ...opts,
    children: Array.isArray(runs) ? runs : [txt(runs)],
  });
}
function centered(runs, opts = {}) {
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { after: 120, line: 360 },
    ...opts,
    children: Array.isArray(runs) ? runs : [txt(runs)],
  });
}
const D = DADOS;
// ========================================
// DOCUMENTO
// ========================================
const doc = new Document({
  styles: {
    default: { document: { run: { font: "Times New Roman", size: 24 } } },
  },
  sections: [{
    properties: {
      page: {
        size: { width: 11906, height: 16838 },
        margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 },
      },
    },
    headers: {
      default: new Header({
        children: [
          new Paragraph({
            alignment: AlignmentType.RIGHT,
            children: [txt("Lopes e Monteiro Advogados Associados", { size: 16, color: "888888", italics: true })],
          }),
        ],
      }),
    },
    footers: {
      default: new Footer({
        children: [
          new Paragraph({
            alignment: AlignmentType.CENTER,
            border: { top: { style: BorderStyle.SINGLE, size: 1, color: "AAAAAA", space: 4 } },
            children: [
              txt("OAB/MG 108.591 | Pag. ", { size: 16, color: "888888" }),
              new TextRun({ children: [PageNumber.CURRENT], font: "Times New Roman", size: 16, color: "888888" }),
              txt("/", { size: 16, color: "888888" }),
              new TextRun({ children: [PageNumber.TOTAL_PAGES], font: "Times New Roman", size: 16, color: "888888" }),
            ],
          }),
        ],
      }),
    },
    children: [
      // ============================================================
      // ENDEREÇAMENTO
      // ============================================================
      centered([
        bold(`EXCELENTISSIMO(A) SENHOR(A) DOUTOR(A) JUIZ(A) DE DIREITO DO JUIZADO ESPECIAL DA FAZENDA PUBLICA DA COMARCA DE ${D.comarca} - MG`),
      ], { spacing: { after: 360 } }),
      emptyLine(),
      // OBSERVAÇÕES LATERAIS
      new Paragraph({
        alignment: AlignmentType.RIGHT,
        spacing: { after: 40 },
        children: [txt("Rito Sumarissimo - Lei 12.153/2009", { italics: true, size: 20, color: "555555" })],
      }),
      new Paragraph({
        alignment: AlignmentType.RIGHT,
        spacing: { after: 300 },
        children: [txt("Juizado Especial da Fazenda Publica", { italics: true, size: 20, color: "555555" })],
      }),
      // TÍTULO
      centered([bold("ACAO DE OBRIGACAO DE FAZER C/C COBRANCA", { size: 26 })]),
      centered([txt(`(Incentivo Financeiro Adicional - IFA - ${D.tipoAgente})`, { italics: true, size: 20, color: "555555" })]),
      emptyLine(),
      // ============================================================
      // QUALIFICAÇÃO DO AUTOR
      // ============================================================
      para([
        bold(D.autorNome),
        txt(`, ${D.autorNacionalidade}, ${D.autorEstadoCivil}, ${D.autorProfissao}, portador(a) do RG n. ${D.autorRG} e inscrito(a) no CPF sob o n. ${D.autorCPF}, residente e domiciliado(a) na ${D.autorEndereco}, Bairro ${D.autorBairro}, ${D.autorCidade}/MG, CEP: ${D.autorCEP}, telefone: ${D.autorTelefone}, e-mail: ${D.autorEmail}, vem, respeitosamente, por meio de seu advogado infra-assinado, propor a presente`),
      ]),
      centered([bold("ACAO DE OBRIGACAO DE FAZER C/C COBRANCA")]),
      // QUALIFICAÇÃO DO RÉU
      para([
        txt("em face do "),
        bold(`MUNICIPIO DE ${D.municipioNome}/MG`),
        txt(`, pessoa juridica de direito publico interno, inscrita no CNPJ sob o n. ${D.municipioCNPJ}, com sede na ${D.municipioEndereco}, Bairro ${D.municipioBairro}, CEP: ${D.municipioCEP}, telefone: ${D.municipioTelefone}, pelos fatos e fundamentos a seguir expostos:`),
      ]),
      emptyLine(),
      // ============================================================
      // I — DOS FATOS
      // ============================================================
      heading("I - DOS FATOS"),
      para([
        txt(`O(A) autor(a) exerce a funcao de `),
        bold(D.tipoAgentePorExtenso),
        txt(` no Municipio de ${D.municipioNome}/MG desde ${D.dataAdmissao}, estando lotado(a) na ${D.lotacao}, com matricula funcional n. ${D.matricula}, percebendo salario base de R$ ${D.salarioBase}.`),
      ]),
      para([
        txt("A Uniao Federal, por meio do Fundo Nacional de Saude (FNS), repassa regularmente ao Municipio reu recursos financeiros destinados ao pagamento do "),
        bold("Incentivo Financeiro Adicional (IFA)"),
        txt(` aos ${D.tipoAgente}s, nos termos do Art. 9.-D da Lei Federal n. 11.350/2006 e da Portaria n. 1.350/GM/2002 do Ministerio da Saude.`),
      ]),
      para([
        txt("Todavia, o Municipio reu, embora receba regularmente tais repasses federais, "),
        bold("nao os transfere ao(a) autor(a)"),
        txt(`, retendo indevidamente valores que possuem destinacao exclusiva e vinculada ao pagamento dos agentes de saude.`),
      ]),
      para([
        txt(`O(A) autor(a) deixou de receber o IFA referente ao periodo de `),
        bold(`${D.periodoInicio} a ${D.periodoFim}`),
        txt(`, totalizando `),
        bold(`${D.mesesDevidos} meses`),
        txt(` sem o devido repasse, causando-lhe significativo prejuizo financeiro.`),
      ]),
      emptyLine(),
      // ============================================================
      // II — DO DIREITO
      // ============================================================
      heading("II - DO DIREITO"),
      // II.1 — Competência
      subHeading("II.1 - Da competencia do Juizado Especial da Fazenda Publica"),
      para([
        txt("A presente acao e da competencia absoluta do Juizado Especial da Fazenda Publica, nos termos do Art. 2., "),
        txt("caput", { italics: true }),
        txt(`, da Lei Federal n. 12.153/2009, que estabelece competencia para processar e julgar causas de ate 60 (sessenta) salarios minimos contra os entes publicos municipais. O valor da presente causa - R$ ${D.valorCausa} - encontra-se dentro do limite legal.`),
      ]),
      para([
        txt("Nos termos do Art. 2., paragrafo 4., da mesma Lei, ficam excluidas da competencia do Juizado as acoes que demandem providencias especiais nao previstas neste rito, o que nao e o caso dos autos, tratando-se de acao de cobranca com pedido de obrigacao de fazer perfeitamente compativel com o rito sumarissimo."),
      ]),
      // II.2 — IFA
      subHeading("II.2 - Do Incentivo Financeiro Adicional (IFA)"),
      para([
        txt("A Constituicao Federal, em seu Art. 198, paragrafo 7., preve que o piso salarial dos Agentes Comunitarios de Saude e dos Agentes de Combate as Endemias sera definido por lei federal, cabendo a Uniao prestar assistencia financeira complementar aos entes da Federacao."),
      ]),
      para([
        txt("A "),
        bold("Lei Federal n. 11.350/2006"),
        txt(", em seu "),
        bold("Art. 9.-D"),
        txt(", estabelece expressamente que os ACS e ACE farao jus ao recebimento do incentivo financeiro adicional, pago pela Uniao de forma complementar ao piso salarial. Tal dispositivo nao confere ao Municipio qualquer discricionariedade quanto ao repasse: trata-se de verba federal, com destinacao especifica e vinculada."),
      ]),
      para([
        txt("A "),
        bold("Portaria n. 1.350/GM/2002"),
        txt(" do Ministerio da Saude regulamenta o repasse desse incentivo, determinando que os valores sejam transferidos fundo a fundo (FNS para Fundo Municipal de Saude), com destinacao "),
        bold("exclusiva"),
        txt(" ao pagamento dos agentes."),
      ]),
      para([
        txt("A "),
        bold("Portaria n. 674/2003"),
        txt(" define que o incentivo adicional corresponde a "),
        bold("1/12 avos (13a parcela)"),
        txt(" do incentivo financeiro anual, devendo ser integralmente repassado ao agente como complemento remuneratorio. O Municipio e mero intermediario dos recursos federais."),
      ]),
      // II.3 — Jurisprudência
      subHeading("II.3 - Da jurisprudencia favoravel"),
      para([
        txt("O Superior Tribunal de Justica, no "),
        bold("AREsp 2.501.949/PA"),
        txt(" (Min. Francisco Falcao, j. 22/04/2024), consolidou o entendimento de que o incentivo adicional constitui "),
        bold("direito subjetivo dos agentes"),
        txt(", cabendo ao Municipio o dever legal de repassar integralmente os valores recebidos da Uniao, sob pena de enriquecimento ilicito."),
      ]),
      para([
        txt("O Tribunal de Justica de Minas Gerais, na "),
        bold("Apelacao 1.0000.25.109792-9/001"),
        txt(" (Des. Luzia Divina de Paula, 3a Camara Civel, DJe 21/10/2025), concedeu "),
        bold("provimento integral"),
        txt(" para condenar o municipio ao repasse retroativo do IFA com correcao monetaria pelo IPCA-E e SELIC, reconhecendo a obrigacao municipal de transferencia."),
      ]),
      para([
        txt("O TJMS, na "),
        bold("APL 0821850-45.2015.8.12.0001"),
        txt(" (Des. Alexandre Bastos, j. 06/02/2019), igualmente reconheceu que o incentivo adicional e devido diretamente aos agentes, condenando o municipio ao pagamento retroativo."),
      ]),
      // II.4 — Enriquecimento sem causa
      subHeading("II.4 - Do enriquecimento sem causa"),
      para([
        txt("Ao reter recursos federais com destinacao especifica e vinculada, o Municipio reu incorre em "),
        bold("enriquecimento sem causa"),
        txt(" (Art. 884 do Codigo Civil), utilizando para fins diversos verbas que pertencem legitimamente ao(a) autor(a). A restituicao e medida que se impoe."),
      ]),
      para([
        txt("Ademais, a conduta municipal viola os principios da "),
        bold("legalidade"),
        txt(" (descumprimento da legislacao federal), "),
        bold("moralidade"),
        txt(" (retencao de verba alimentar de trabalhador essencial) e "),
        bold("eficiencia"),
        txt(" (Art. 37, "),
        txt("caput", { italics: true }),
        txt(", CF)."),
      ]),
      emptyLine(),
      // ============================================================
      // III — DO CÁLCULO
      // ============================================================
      heading("III - DO CALCULO DO VALOR DEVIDO"),
      para([
        txt(`O(A) autor(a) faz jus ao valor mensal estimado de `),
        bold(`R$ ${D.valorMensal}`),
        txt(` a titulo de IFA, conforme dados dos repasses do Fundo Nacional de Saude ao Municipio reu e o numero de ${D.tipoAgente}s cadastrados no SCNES.`),
      ]),
      para([
        txt(`Considerando o periodo de ${D.mesesDevidos} meses sem repasse (${D.periodoInicio} a ${D.periodoFim}), o valor total devido, antes da correcao monetaria e juros, e de `),
        bold(`R$ ${D.valorTotal}`),
        txt("."),
      ]),
      para([
        txt("Os valores devem ser corrigidos monetariamente pelo "),
        bold("IPCA-E"),
        txt(" ate dezembro/2021, e pela taxa "),
        bold("SELIC"),
        txt(" a partir de janeiro/2022, nos termos da "),
        bold("EC 113/2021, Art. 3."),
        txt(". Os juros moratorios incidem a partir da citacao valida (Art. 405, CC c/c Art. 240, CPC), conforme RE 870.947 (Tema 810, STF) e REsp 1.495.146 (Tema 905, STJ)."),
      ]),
      emptyLine(),
      // ============================================================
      // IV — DA TUTELA DE URGÊNCIA
      // ============================================================
      heading("IV - DA TUTELA DE URGENCIA"),
      para([
        txt("Nos termos do Art. 300 do CPC, requer-se a concessao de tutela de urgencia para determinar ao Municipio reu o imediato repasse do IFA ao(a) autor(a), tendo em vista:"),
      ]),
      para([
        bold("a) Probabilidade do direito: "),
        txt("demonstrada pela legislacao federal, portarias ministeriais e jurisprudencia consolidada do STJ e TJMG;"),
      ]),
      para([
        bold("b) Perigo de dano irreparavel: "),
        txt("o IFA possui natureza alimentar. O(A) autor(a), profissional essencial ao SUS, esta sendo privado(a) de verba remuneratoria a que faz jus, comprometendo seu sustento e de sua familia;"),
      ]),
      para([
        bold("c) Reversibilidade: "),
        txt("a medida e perfeitamente reversivel, tratando-se de obrigacao pecuniaria contra ente publico solvente."),
      ]),
      emptyLine(),
      // ============================================================
      // V — DOS PEDIDOS
      // ============================================================
      heading("V - DOS PEDIDOS"),
      para([txt("Ante o exposto, requer-se a Vossa Excelencia:")]),
      para([
        bold("a) "),
        txt("A citacao do Municipio reu para, querendo, apresentar contestacao na audiencia de conciliacao;"),
      ]),
      para([
        bold("b) "),
        txt("A concessao de "),
        bold("TUTELA DE URGENCIA"),
        txt(" (Art. 300, CPC) para determinar o imediato repasse mensal do IFA ao(a) autor(a), sob pena de multa diaria;"),
      ]),
      para([
        bold("c) "),
        txt("No merito, a "),
        bold("TOTAL PROCEDENCIA"),
        txt(" da acao, condenando o Municipio reu a:"),
      ]),
      para([
        txt("(i) cumprir a "),
        bold("obrigacao de fazer"),
        txt(`, repassando mensalmente o IFA ao(a) autor(a), na qualidade de ${D.tipoAgentePorExtenso};`),
      ]),
      para([
        txt("(ii) pagar os "),
        bold("valores retroativos"),
        txt(` devidos no periodo de ${D.periodoInicio} a ${D.periodoFim} (${D.mesesDevidos} meses), no montante de R$ ${D.valorTotal}, corrigidos pelo IPCA-E/SELIC, acrescidos de juros moratorios desde a citacao;`),
      ]),
      para([
        bold("d) "),
        txt("A condenacao do reu ao pagamento de "),
        bold("honorarios advocaticios"),
        txt(", nos termos do Art. 55 da Lei 9.099/95 c/c Art. 27 da Lei 12.153/2009;"),
      ]),
      para([
        bold("e) "),
        txt("A producao de todas as provas admitidas em direito, especialmente documental, e requisicao de informacoes ao Fundo Nacional de Saude sobre os repasses ao Municipio reu."),
      ]),
      emptyLine(),
      // VALOR DA CAUSA
      para([
        txt("Da-se a causa o valor de "),
        bold(`R$ ${D.valorCausa}`),
        txt("."),
      ]),
      emptyLine(),
      // ENCERRAMENTO
      centered([txt("Nestes termos, pede deferimento.")]),
      emptyLine(),
      centered([txt(`${D.autorCidade}/MG, data da assinatura eletronica.`)]),
      emptyLine(),
      emptyLine(),
      centered([bold("Eldbrendo Pereira Monteiro")]),
      centered([txt("OAB/MG 108.591")]),
      centered([txt("Lopes e Monteiro Advogados Associados")]),
    ],
  }],
});
// ========================================
// GERAR ARQUIVO
// ========================================
const nomeArquivo = `IFA_JE_${D.tipoAgente}_${D.autorNome.replace(/\s+/g, "_").substring(0, 30)}.docx`;
Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync(nomeArquivo, buffer);
  console.log(`\n  Peticao gerada com sucesso!`);
  console.log(`  Arquivo: ${nomeArquivo}`);
  console.log(`  Tamanho: ${(buffer.length / 1024).toFixed(0)} KB`);
  console.log(`  Autor(a): ${D.autorNome}`);
  console.log(`  Tipo: ${D.tipoAgente} - ${D.tipoAgentePorExtenso}`);
  console.log(`  Municipio: ${D.municipioNome}/MG`);
  console.log(`  Valor: R$ ${D.valorTotal} (${D.mesesDevidos} meses)\n`);
});
