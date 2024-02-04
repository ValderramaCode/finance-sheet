function shceduleBills() {
  const spreadSheet = SpreadsheetApp.getActiveSheet();

  const calendarioId = spreadSheet.getRangeByName('calendario');
  const mes = spreadSheet.getRangeByName('mes');
  const ano = spreadSheet.getRangeByName('ano');

  const contas2 = spreadSheet.getRange('contas');
  const contas = spreadSheet.getRangeByName('contas');
  const empresas = spreadSheet.getRangeByName('empresas');
  const valores = spreadSheet.getRangeByName('valores');
  const vencimentos = spreadSheet.getRangeByName('vencimentos');

  if (!calendarioId || !mes || !ano) {
    throw new Error(`Intervalos nomeados DE CONFIGURAÇÃO ainda não criados 
        totalmente. Favor acessar no menu "Dados -> Intervalos Nomeados"
        e criar os intervalos de CONFIGURAÇÃO "mes", "ano" e "calendario",
        com as respectivas colunas da tabela`);
  }
  if (!contas || !empresas || !valores || !vencimentos) {
    throw new Error(`Intervalos nomeados DE LANÇAMENTO ainda não criados 
        totalmente. Favor acessar no menu "Dados -> Intervalos Nomeados"
        e criar os intervalos "contas", "empresas", "valores" e 
        "vencimentos", com as respectivas colunas da tabela`);
  }

  if (calendarioId.length === 0 || mes.length === 0 || ano.length === 0) {
    throw new Error(`Intervalos nomeados DE CONFIGURAÇÃO criados, porém não há VALORES CONFIGURADOS.
        Por favor insira os DADOS DE CONFIGURAÇÃO propriamente para execução do programa.`);
  }

  if (contas.length === 0 || empresas.length === 0 || valores.length === 0 || vencimentos.length === 0) {
    throw new Error(`Intervalos nomeados DE LANÇAMENTO criados, porém não há VALORES PARA LANÇAMENTO.
        Por favor insira os DADOS PARA LANÇAMENTO propriamente para execução do programa.`);
  }
}

export { shceduleBills };
