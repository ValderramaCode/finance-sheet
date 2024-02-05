/* eslint-disable no-plusplus */
/* eslint-disable no-continue */

function obterPrimeiroDiaUtilAnteriorEmMinutos(data) {
  // Configuração do fuso horário para o Brasil
  const timeZone = 'America/Sao_Paulo';

  // Obtém a data no fuso horário desejado
  const dataBrasil = new Date(data.toLocaleString(undefined, { timeZone }));

  // Verifica se é final de semana (sábado ou domingo)
  const diaDaSemana = dataBrasil.getDay();
  if (diaDaSemana === 0) {
    // Se for domingo, subtrai 2 dias para obter sexta-feira
    dataBrasil.setDate(dataBrasil.getDate() - 2);
  } else if (diaDaSemana === 6) {
    // Se for sábado, subtrai 1 dia para obter sexta-feira
    dataBrasil.setDate(dataBrasil.getDate() - 1);
  } else {
    // Para os demais dias da semana, subtrai um dia
    dataBrasil.setDate(dataBrasil.getDate() - 1);
  }

  // Converte a data para minutos e retorna o resultado
  const minutos = dataBrasil.getTime() / 60000;
  return minutos;
}

function dataProximoDiaUtil(data) {
  // Configuração do fuso horário para o Brasil
  const timeZone = 'America/Sao_Paulo';

  // Obtém a data no fuso horário desejado
  const dataBrasil = new Date(data.toLocaleString(undefined, { timeZone }));

  // Verifica se é final de semana (sábado ou domingo)
  const diaDaSemana = dataBrasil.getDay();

  if (diaDaSemana === 0) {
    // Se for domingo, avança para segunda-feira
    dataBrasil.setDate(dataBrasil.getDate() + 1);
  } else if (diaDaSemana === 6) {
    // Se for sábado, avança para segunda-feira
    dataBrasil.setDate(dataBrasil.getDate() + 2);
  }

  // Retorna a data do próximo dia útil
  return dataBrasil;
}

function scheduleBills() {
  const ui = SpreadsheetApp.getUi();
  const spreadSheet = SpreadsheetApp.getActiveSheet();

  const calendarioId = spreadSheet.getRange('C2').getValue();
  const mes = spreadSheet.getRange('A2').getValue();
  const ano = spreadSheet.getRange('B2').getValue();

  const contasRange = spreadSheet.getRange('A6:A33');
  const empresasRange = spreadSheet.getRange('B6:B33');
  const valoresRange = spreadSheet.getRange('C6:C33');
  const vencimentosRange = spreadSheet.getRange('E6:E33');

  if (!calendarioId || !mes || !ano) {
    const mensagemErro = `Dados DE CONFIGURAÇÃO ainda não criados 
        totalmente. Favor preencher os dados de CONFIGURAÇÃO "mês",
        "ano" e "calendario", nas respectivas colunas da tabela.`;
    SpreadsheetApp.getUi().alert('Configuração Incompleta', mensagemErro, ui.ButtonSet.OK);
    throw new Error(mensagemErro);
  }
  if (!contasRange || !empresasRange || !valoresRange || !vencimentosRange) {
    const mensagemErro = `Dados DE LANÇAMENTO ainda não criados 
      totalmente. Favor preencher os dados de LANÇAMENTO "contas",
      "empresas", "quantia" e "vencimentos", nas respectivas 
      colunas da tabela`;
    SpreadsheetApp.getUi().alert('Lançamento Incompleto', mensagemErro, ui.ButtonSet.OK);
    throw new Error(`mensagemErro`);
  }

  // if (calendarioId.length === 0 || mes.length === 0 || ano.length === 0) {
  //   throw new Error(`Intervalos nomeados DE CONFIGURAÇÃO criados, porém não há VALORES CONFIGURADOS.
  //       Por favor insira os DADOS DE CONFIGURAÇÃO propriamente para execução do programa.`);
  // }

  const corDeBloqueio = '#ff9900';
  const indicesLinhasInvalidas = new Array(contasRange.getNumRows()).fill(0);

  const coresDeFundoDasLinhas = contasRange.getBackgrounds().flatMap(([bg], index) => bg);

  const contasFlat = contasRange.getValues().flatMap(([conta]) => String(conta).trim());

  const empresasFlat = empresasRange.getValues().flatMap(([empresa]) => String(empresa).trim());

  const valoresFlat = valoresRange.getValues().flatMap(([valor]) => valor);

  const vencimentosFlat = vencimentosRange.getValues().flatMap(([vencimento]) => vencimento);

  const numeroDeLinhas = contasRange.getNumRows();
  for (let indexLinha = 0; indexLinha < numeroDeLinhas; indexLinha++) {
    const isLinhaBloqueada = coresDeFundoDasLinhas[indexLinha] === corDeBloqueio;
    if (isLinhaBloqueada) {
      indicesLinhasInvalidas[indexLinha] = 1;
      continue;
    }

    const isContaValida = contasFlat[indexLinha] && contasFlat[indexLinha] !== '-';
    if (!isContaValida) {
      indicesLinhasInvalidas[indexLinha] = 1;
      continue;
    }

    const isValorValido =
      valoresFlat[indexLinha] &&
      vencimentosFlat[indexLinha] !== '?' &&
      vencimentosFlat[indexLinha] !== '??' &&
      Number(valoresFlat[indexLinha]) > 0;
    if (!isValorValido) {
      indicesLinhasInvalidas[indexLinha] = 1;
      continue;
    }

    const isVencimentoValido =
      vencimentosFlat[indexLinha] &&
      vencimentosFlat[indexLinha] !== '?' &&
      vencimentosFlat[indexLinha] !== '??' &&
      Number(vencimentosFlat[indexLinha] > 0);
    if (!isVencimentoValido) {
      indicesLinhasInvalidas[indexLinha] = 1;
      continue;
    }
  }

  // TODO: abstrair esse metodo de conseguir contas validas.
  const contasValidas = contasFlat.filter((_, indice) => !indicesLinhasInvalidas[indice]);

  const empresasValidas = empresasFlat.filter((_, indice) => !indicesLinhasInvalidas[indice]);

  const valoresValidos = valoresFlat.filter((_, indice) => !indicesLinhasInvalidas[indice]);

  const vencimentosValidos = vencimentosFlat.filter((_, indice) => !indicesLinhasInvalidas[indice]);

  if (
    contasValidas.length === 0 ||
    empresasValidas.length === 0 ||
    valoresValidos.length === 0 ||
    vencimentosValidos.length === 0
  ) {
    const mensagemErro = `Dados DE LANÇAMENTO criados,
      porém não há VALORES PARA LANÇAMENTO. Por favor insira
      os DADOS PARA LANÇAMENTO propriamente para execução 
      do programa.`;
    SpreadsheetApp.getUi().alert('Lançamento Incompleto', mensagemErro, ui.ButtonSet.OK);
    throw new Error(mensagemErro);
  }

  // Marca evento no proximo dia util
  const contasCalendario = CalendarApp.getCalendarById(calendarioId);
  // Código para DEBUG
  // const datas = [];
  const numeroDeLinhasValidas = contasValidas.length;
  for (let indexLinha = 0; indexLinha < numeroDeLinhasValidas; indexLinha++) {
    const tituloConta = `${contasValidas[indexLinha]} 
    ${empresasValidas[indexLinha] && empresasValidas[indexLinha] !== '-' ? `[${empresasValidas[indexLinha]}]` : ''}`;

    const diaVencimento = vencimentosValidos[indexLinha];
    const dataVencimento = new Date(ano, mes - 1, diaVencimento);
    const dataUtilVencimento = dataProximoDiaUtil(dataVencimento);
    const dataUtilVencimentoInicial = new Date(ano, mes - 1, dataUtilVencimento, 16, 30);
    const dataUtilVencimentoFinal = new Date(ano, mes - 1, dataUtilVencimento, 17, 0);
    const eventoConta = contasCalendario.createEvent(tituloConta, dataUtilVencimentoInicial, dataUtilVencimentoFinal);

    const primeiroDiaUtilAnteriorEmMinutos = obterPrimeiroDiaUtilAnteriorEmMinutos(dataVencimento);
    const minutosAntesDesdeAs9 = 450;
    const minutosAntesDesdeAs10 = 390;
    const minutosAntesDesdeAs12 = 270;
    const minutosAntesDesdeAs15 = 90;
    eventoConta.addPopupReminder(primeiroDiaUtilAnteriorEmMinutos);
    eventoConta.addPopupReminder(minutosAntesDesdeAs9);
    eventoConta.addPopupReminder(minutosAntesDesdeAs10);
    eventoConta.addPopupReminder(minutosAntesDesdeAs12);
    eventoConta.addPopupReminder(minutosAntesDesdeAs15);
    // Código para DEBUG
    // datas.push({
    //   tituloConta,
    //   diaVencimento,
    //   dataVencimento,
    //   dataUtilVencimento,
    //   primeiroDiaUtilAnteriorEmMinutos
    // });
  }
  ui.alert(
    'Lançamentos feitos com sucesso!',
    'Aguarde alguns minutos enquanto ocorre a sincronização com o calendário em seu APP.',
    ui.ButtonSet.OK
  );
}

function confirmAction() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    'Confirmar Lançamentos',
    'Deseja realmente agendar esses lançamentos?',
    ui.ButtonSet.YES_NO
  );
  if (response === ui.Button.YES) {
    scheduleBills();
  }
}

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Planilha de Finanças').addItem('Agendar vencimentos', 'confirmAction').addToUi();
}

export { scheduleBills, confirmAction, onOpen, obterPrimeiroDiaUtilAnteriorEmMinutos, dataProximoDiaUtil };
