function getDadosCircle() {
    const POS_STATUS = 5;
    const POS_MATRICULA = 0;
    const agregador_status = {};
    
    const planilha = SpreadsheetApp.getActiveSpreadsheet();
    const aba = planilha.getSheetByName("Dados Academicos");
    const dados_puros = aba.getRange(2, 1, aba.getLastRow(),aba.getLastColumn()).getDisplayValues();
    
    for (let i=0; i < dados_puros.length; i++)
    {
      const status = dados_puros[i][POS_STATUS];
      if (!status) continue;
      agregador_status[status] = (agregador_status[status] || 0) + 1;
    }

    const chartData = [["Status Matrícula", "Quantidade"]];

    for (const status in agregador_status) {
      chartData.push([status, agregador_status[status]]);
    }
    console.log(chartData)
    return chartData;
}

function getDadosLine() {
  const planilha = SpreadsheetApp.getActiveSpreadsheet();
  const aba = planilha.getSheetByName("Disciplinas");

  const readRange = aba
    .getRange(2, 1, aba.getLastRow() - 1, aba.getLastColumn())
    .getDisplayValues();

  // Monta a tabela começando pelo cabeçalho correto
  const chartData = [
    ["Disciplina", "Qnt Vagas", "Qnt Inscritos"]
  ];

  for (const r of readRange) {
    const disciplina = r[1];
    const vagas = parseInt(r[5]);
    const inscritos = parseInt(r[6]);

    // Ignora linhas vazias ou inválidas
    if (!disciplina || disciplina.trim() === "") continue;
    if (isNaN(vagas)) continue;
    if (isNaN(inscritos)) continue;

    chartData.push([disciplina, vagas, inscritos]);
  }
  console.log(chartData);
  return chartData;
}

function getDadosCRMedio() {
    const POS_CR = 4;
    const POS_MATRICULA = 0;
    var agregador_cr = 0;
    var qnt = 0;
    
    const planilha = SpreadsheetApp.getActiveSpreadsheet();
    const aba = planilha.getSheetByName("Dados Academicos");
    const dados_puros = aba.getRange(2, 1, aba.getLastRow(),aba.getLastColumn()).getDisplayValues();
    
    for (let i=0; i < dados_puros.length; i++)
    {
      const cr = Number(dados_puros[i][POS_CR]);
      if (!cr) continue;
      qnt++;
      agregador_cr += cr;
    }

    const media = agregador_cr/qnt;

    const chartData = [["Meta","Média"]];
    chartData.push([10, media]);

    console.log(chartData)
    return chartData;
}


function getDadosAno() {
    const POS_ANO = 2;
    const POS_MATRICULA = 0;
    var agregador_ano = {};
    
    const planilha = SpreadsheetApp.getActiveSpreadsheet();
    const aba = planilha.getSheetByName("Dados Aluno");
    const dados_puros = aba.getRange(2, 1, aba.getLastRow(),aba.getLastColumn()).getDisplayValues();
    
    for (let i=0; i < dados_puros.length; i++)
      {
        const ano = dados_puros[i][POS_ANO];
        if (!ano) continue;
        agregador_ano[ano] = (agregador_ano[ano] || 0) + 1;
      }

    const chartData = [["Ano", "Quantidade"]];

    for (const ano in agregador_ano) {
      chartData.push([ano, agregador_ano[ano]]);
    }
    console.log(chartData)
    return chartData;
}

function getAprovacao() {
  const planilha = SpreadsheetApp.getActiveSpreadsheet();
  const aba1 = planilha.getSheetByName("Disciplinas");
  const chartData = [['Disciplina', 'Porcentagem Aprovação','Nº Alunos']];

  const readRange1 = aba1
    .getRange(2, 1, aba1.getLastRow() - 1, aba1.getLastColumn())
    .getDisplayValues();

  // Monta a tabela começando pelo cabeçalho correto
  const consulta1 = {}

  for (const r of readRange1) {
    const disciplina = r[1];
    const vagas = parseInt(r[5]);
    const inscritos = parseInt(r[6]);

    // Ignora linhas vazias ou inválidas
    if (!disciplina || disciplina.trim() === "") continue;
    if (isNaN(vagas)) continue;
    if (isNaN(inscritos)) continue;

    consulta1[disciplina]= [vagas, inscritos];
  }
  console.log(consulta1);

  const aba2 = planilha.getSheetByName("Dados do Periodo");

  const readRange2 = aba2
    .getRange(2, 1, aba2.getLastRow() - 1, aba2.getLastColumn())
    .getDisplayValues();

    const consulta2 = {}
    for (const r of readRange2) {
    const disciplina = r[0];
    const resultado = (r[5] === "aprovado")? 1:0;

    // Ignora linhas vazias ou inválidas
    if (!disciplina || disciplina.trim() === "") continue;
    if (!consulta1[disciplina]) continue;
    if (isNaN(resultado)) continue;


    consulta2[disciplina] = (consulta2[disciplina] || 0) + resultado ;
  }

  for (const disciplina in consulta2) {
    chartData.push([disciplina, consulta2[disciplina]/consulta1[disciplina][1],consulta1[disciplina][1]]);
  }

  console.log(chartData)

  return chartData;
}

