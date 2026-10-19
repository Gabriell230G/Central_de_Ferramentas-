using System;
using System.Data;
using System.Windows.Forms;
using System.Data.OleDb;
using System.Threading;

namespace ComplementaCadastro
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        public static void gravaLOG(string str_log)
        {
            string str_sql = "INSERT INTO TB_LOG ( ID_Usuario, Log, DataHora) VALUES ('" + Environment.UserName + "', '" + str_log + "' ,'" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "')";
            Conexao.executar(str_sql);
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            preencheGrade("Bem vindo ao Sistema Atualiza Envolvidos!");

            carregaGridDados();
            if (grid_dados.Rows.Count > 0)
            {
                preencheGrade("Foram encontrados registros do último processamento no Banco de Dados");
                preencheGrade("Para continuar Conecte ao GCPJ, para iniciar novo Processamento clique em Importar Planilha");
                preencheGrade("Certifique-se de que a planilha esteja no diretório do sistema e com as seguintes colunas: BRADESCO | ENVOLVIDO | CPF_CNPJ | FILIAL | DIGITO");

                //Verifica se tem distribuição em mais de uma máquina
                string sql = "SELECT TB_ENVOLVIDOS.MAQUINA FROM TB_ENVOLVIDOS GROUP BY TB_ENVOLVIDOS.MAQUINA;";
                DataTable dt = Conexao.GeraTBL(sql);
                if (dt.Rows.Count > 1)
                {
                    txtMaquina.Text = "";
                }
            }
            else
            {
                preencheGrade("Importe a planilha para iniciar o processamento.");
                preencheGrade("Certifique-se de que a planilha esteja no diretório do sistema e com as seguintes colunas: BRADESCO | ENVOLVIDO | CPF_CNPJ | FILIAL | DIGITO");
            }
            gravaLOG("ACESSOU O SISTEMA");
        }

        private void btn_Importar_Click(object sender, EventArgs e)
        {
            if (grid_dados.Rows.Count > 0)
            {
                if (MessageBox.Show("Deseja iniciar uma nova importação?\nOs dados já importados serão perdidos.", "Atenção", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1) != DialogResult.Yes)
                {
                    return;
                }
            }

            Conexao.executarProcedure("000_LIMPAR_TB_ENVOLVIDOS");

            if (Conexao.executarProcedure("001_ACRESCIMO_TB_ENVOLVIDOS") == false)
            {
                MessageBox.Show("Verifique se o arquivo possui as seguintes colunas:  BRADESCO | ENVOLVIDO | CPF_CNPJ | FILIAL | DIGITO ", "Erro ao importar o arquivo", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            carregaGridDados();
            gravaLOG("IMPORTOU PLANILHA");
            MessageBox.Show("Importação efetuada com sucesso!", "Pronto", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void carregaGridDados()
        {
            string sql = "SELECT TB_ENVOLVIDOS.CNRO_PROCS_BDSCO as GCPJ, TB_ENVOLVIDOS.NOME, TB_ENVOLVIDOS.[CPF/CNPJ], TB_ENVOLVIDOS.[FILIAL (PJ)], TB_ENVOLVIDOS.DIGITO, TB_ENVOLVIDOS.ATUALIZADO, TB_ENVOLVIDOS.OBS, TB_ENVOLVIDOS.MAQUINA FROM TB_ENVOLVIDOS;";
            grid_dados.DataSource = Conexao.GeraTBL(sql);
            if (grid_dados.Rows.Count > 0)
            {
                StatusLabel_Registros.Text = "Registro: " + (grid_dados.CurrentRow.Index + 1) + " de " + grid_dados.Rows.Count;
            }
            else
            {
                StatusLabel_Registros.Text = "";
            }
        }

        private void btnConectar_Click(object sender, EventArgs e)
        {
            if (Automacao.localizarJanela("GCPJ") == false)
            {
                Automacao.novaJanela("https://intranet85.net.bradesco.com.br/gcpj/");
                preencheGrade("Abrindo nova instância do GCPJ, por favor, insira seu usuário e senha");
            }

            preencheGrade("GCPJ localizado, por favor, inicie o processamento...");
        }

        private void preencheGrade(String texto)
        {
            grade.Items.Add(DateTime.Now.ToShortTimeString() + ' ' + texto);
        }
                 
        private void btnDistribuir_Click(object sender, EventArgs e)
        {
            if (txtMaquina.Text.ToString().Length == 0 || txtMaquina.Text.Equals(""))
            {
                //consistencia de preenchimento do campo
                MessageBox.Show("Por favor, preencha um número de maquinas para distribuir!", "Atenção", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            int qtdeMaquina = 0;

            try
            {
                qtdeMaquina = int.Parse(txtMaquina.Text);
            }
            catch (Exception)
            {
                MessageBox.Show("O número de maquinas deve ser um número", "Atenção", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }


            preencheGrade("Distribuindo maquinas...");

            int contador = 1;
            OleDbDataReader dr = Conexao.consultar("SELECT * FROM TB_ENVOLVIDOS WHERE ATUALIZADO = FALSE");

            while (dr.Read())
            {
                Conexao.executar("UPDATE TB_ENVOLVIDOS SET MAQUINA = " + contador + " WHERE CNRO_PROCS_BDSCO = " + dr.GetDouble(0));
                contador++;
                if (contador > qtdeMaquina)
                {
                    contador = 1;
                }
            }
            gravaLOG("EFETUOU DISTRIBUÍÇÃO DE PROCESSAMENTO EM " + qtdeMaquina  + " MÁQUINAS");
            preencheGrade("Maquinas distribuidas com sucesso!");
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;
            Microsoft.Office.Interop.Excel.Application XcelApp = new Microsoft.Office.Interop.Excel.Application();
            if (grid_dados.Rows.Count > 0)
            {
                XcelApp.Application.Workbooks.Add(Type.Missing);
                for (int i = 1; i < grid_dados.Columns.Count + 1; i++)
                {
                    if (grid_dados.Columns[i - 1].Visible == true)
                    {
                        XcelApp.Cells[1, i] = grid_dados.Columns[i - 1].HeaderText;
                        XcelApp.Cells[1, i].Font.Bold = true;
                    }
                }
                //
                for (int i = 0; i < grid_dados.Rows.Count ; i++)
                {
                    for (int j = 0; j < grid_dados.Columns.Count; j++)
                    {
                        if (grid_dados.Columns[j].Visible == true)
                        {
                            if (grid_dados.Rows[i].Cells[j].Value == null)
                            {
                                XcelApp.Cells[i + 2, j + 1] = ""; 
                            }
                            else if (grid_dados.Rows[i].Cells[j].Value.ToString() == "True")
                            {
                                XcelApp.Cells[i + 2, j + 1] = "SIM";
                            }
                            else if (grid_dados.Rows[i].Cells[j].Value.ToString() == "False")
                            {
                                XcelApp.Cells[i + 2, j + 1] = "NÃO";
                            }
                            else
                            {
                                XcelApp.Cells[i + 2, j + 1] = grid_dados.Rows[i].Cells[j].Value.ToString();
                            }
                        }
                    }
                }
                XcelApp.Columns.AutoFit();
                XcelApp.Visible = true;
            }

            gravaLOG("EXPORTOU GRID EXCEL");
            Cursor.Current = Cursors.Default;
        }

        private void grid_dados_CellMouseUp(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (grid_dados.Rows.Count > 0)
            { StatusLabel_Registros.Text = "Registro: " + (grid_dados.CurrentRow.Index + 1) + " de " + grid_dados.Rows.Count; }
        }

        private void grid_dados_KeyUp(object sender, KeyEventArgs e)
        {
            if (grid_dados.Rows.Count > 0)
            { StatusLabel_Registros.Text = "Registro: " + (grid_dados.CurrentRow.Index + 1) + " de " + grid_dados.Rows.Count; }
        }

        private bool IniciaProcessamento()
        {
            if (Automacao.localizarJanela("GCPJ") == false)
            {
                preencheGrade("ATENÇÃO: Conecte ao GCPJ para continuar.");
                return false;
            }

            if (txtMaquina.Text.ToString().Length == 0 || txtMaquina.Text.Equals(""))
            {
                MessageBox.Show("Por favor, preencha um número de máquinas para distribuir!", "Atenção", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return false;
            }

            int qtdeMaquina = 0;
            try
            {
                qtdeMaquina = int.Parse(txtMaquina.Text);
            }
            catch (Exception)
            {
                MessageBox.Show("O campo Máquinas deve ser um número", "Atenção", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return false;
            }

            return true;
        }

        //ATUALIZAR CPF/CNPJ
        private void btnAtualizaCPF_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Confirma a ação: ATUALIZAR CPF/CNPJ DOS ENVOLVIDOS contidos nesse arquivo importado?", "Atenção", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1) != DialogResult.Yes)
            {
                return;
            }
            if (IniciaProcessamento() == false) { return; }

            int qtdeMaquina = int.Parse(txtMaquina.Text);
            gravaLOG("INICIOU ATUALIZAÇÃO CPF/CNPJ ENVOLVIDO - MÁQUINA " + qtdeMaquina);
            
            string idEnvolvido;
            string tipoEnvolvido;

            string NomeEnvolvSemAcento;
            String nomeEnvolvidLocalizado;

            int qtdOperacoesS5 = 0; // [S5] contador de registros processados nesta execução
            OleDbDataReader dr;
            dr = Conexao.consultar("SELECT * FROM TB_ENVOLVIDOS WHERE ATUALIZADO = FALSE AND MAQUINA = " + qtdeMaquina);

            while (dr.Read())
            {
                // [S5] Limitador de operações: para o processamento ao atingir o limite (ver LimitadorS5.cs)
                if (LimitadorS5.LimiteAtingido(qtdOperacoesS5))
                {
                    preencheGrade(LimitadorS5.MensagemLog());
                    gravaLOG("CONTROLE S5 - LIMITE DE OPERACOES ATINGIDO - ATUALIZAR CPF/CNPJ - MÁQUINA " + qtdeMaquina);
                    break;
                }
                qtdOperacoesS5++;

                if (DateTime.Now.Hour.ToString() == "23")
                {
                    //vamos aguardar o outro dia
                    while (DateTime.Now.Hour.ToString() != "06")
                    {
                        //vamos aguardar ate as 06 da manha.
                        Thread.Sleep(3000000);
                    }
                }

                try
                {
                    Thread.Sleep(800);
                    preencheGrade("Atualizando o envolvido " + dr.GetString(1));
                    Automacao.navegar("https://intranet85.net.bradesco.com.br/gcpj/servlet/cadMantemCadProcAlteracaoAction?acao=carregar&sessao=limpar&estimulo=alteracaoDePara");
                    Automacao.espera();

                    Automacao.setaValorNome("nrProcessoBradesco", dr.GetDouble(0).ToString());
                    Automacao.clicar("btoPesquisar");

                    if (Automacao.mataPossivelMsgBox())
                    {
                        Conexao.executar("UPDATE TB_ENVOLVIDOS SET ATUALIZADO = TRUE, OBS = 'ERRO AO ACESSAR O PROCESSO' WHERE [CPF/CNPJ] = " + dr.GetDouble(2) + "and CNRO_PROCS_BDSCO = " + dr.GetDouble(0));
                        continue;
                    }

                    Automacao.espera();

                    NomeEnvolvSemAcento = ObterStringSemAcentosECaracteresEspeciais(dr.GetString(1));
                    idEnvolvido = Automacao.selecionarTabelaEnvolvido("tab1", dr.GetString(1).Trim(), NomeEnvolvSemAcento);

                    if (idEnvolvido == "ERRO")
                    {
                        Conexao.executar("UPDATE TB_ENVOLVIDOS SET ATUALIZADO = TRUE, OBS = 'ERRO AO LOCALIZAR O NOME DO ENVOLVIDO' WHERE [CPF/CNPJ] = " + dr.GetDouble(2) + "and CNRO_PROCS_BDSCO = " + dr.GetDouble(0));
                        continue;
                    }
                    else if (idEnvolvido == "")
                    {
                        Conexao.executar("UPDATE TB_ENVOLVIDOS SET ATUALIZADO = TRUE, OBS = 'NOME DO ENVOLVIDO NÃO LOCALIZADO' WHERE [CPF/CNPJ] = " + dr.GetDouble(2) + "and CNRO_PROCS_BDSCO = " + dr.GetDouble(0));
                        continue;
                    }
                    else
                    {
                        //Pega o tipo envolvido
                        tipoEnvolvido = idEnvolvido.Substring(idEnvolvido.Length - 1);

                        //Retira o tipo envolvido do id
                        idEnvolvido = idEnvolvido.Substring(0, idEnvolvido.Length - 1);

                        //Verifica se já não está com o mesmo CPF/CNPJ
                        if (idEnvolvido.Length > 9)
                        {
                            if (dr.GetDouble(2).ToString("000000000") + dr.GetDouble(4).ToString("00") == idEnvolvido.Substring(8).Trim() || dr.GetDouble(2).ToString("000000000") + dr.GetDouble(3).ToString("0000") + dr.GetDouble(4).ToString("00") == idEnvolvido.Substring(8).Trim() || dr.GetDouble(2).ToString("00000000") + dr.GetDouble(3).ToString("0000") + dr.GetDouble(4).ToString("00") == idEnvolvido.Substring(8).Trim())
                            {
                                Conexao.executar("UPDATE TB_ENVOLVIDOS SET ATUALIZADO = TRUE, OBS = 'CPF/CNPJ JÁ CADASTRADO PARA O ENVOLVIDO INFORMADO' WHERE [CPF/CNPJ] = " + dr.GetDouble(2) + "and CNRO_PROCS_BDSCO = " + dr.GetDouble(0));
                                continue;
                            }
                        }
                        //  Conexao.executar("UPDATE TB_ENVOLVIDOS SET CODIGO_PESSOA_EXTERNA = " + idEnvolvido + " WHERE [CPF/CNPJ] = " + dr.GetDouble(2) + "and CNRO_PROCS_BDSCO = " + dr.GetDouble(0));
                    }

                    Automacao.clicar("btoEmpresaReu");
                    Automacao.espera();

                    //Verifica se é CPF ou CNPJ
                    if (dr.GetDouble(3).Equals(0))
                    {
                        if (tipoEnvolvido == "A")
                        {
                            Automacao.selecionaCombo("cdCmbTipoEnvolvidoAutor", "3", "value");
                            Automacao.espera();
                            Automacao.ExecutarJS("alterarTipoEnvolvido()");
                            Automacao.espera();
                            Automacao.setaValorNome("cdCodigoAutorDestino", dr.GetDouble(2).ToString("000000000"));
                            Automacao.setaValorNome("cdCtrlAutorDestino", dr.GetDouble(4).ToString("00"));
                            Automacao.ExecutarJS("obterFuncionarioPorCPF('/gcpj')");
                        }
                        else
                        {
                            Automacao.selecionaCombo("cdCmbTipoEnvolvidoReu", "3", "value");
                            Automacao.espera();
                            Automacao.ExecutarJS("alterarTipoReu()");
                            Automacao.espera();
                            Automacao.setaValorNome("cdCodigoReuDestino", dr.GetDouble(2).ToString("000000000"));
                            Automacao.setaValorNome("cdCtrlReuDestino", dr.GetDouble(4).ToString("00"));
                            Automacao.ExecutarJS("obterReuFuncionarioPorCPF('/gcpj')");
                        }


                    }
                    else
                    {
                        if (tipoEnvolvido == "A")
                        {
                            Automacao.selecionaCombo("cdCmbTipoEnvolvidoAutor", "5", "value");
                            Automacao.espera();
                            Automacao.ExecutarJS("alterarTipoEnvolvido()");
                            Automacao.espera();
                            Automacao.setaValorNome("cdCodigoAutorDestino", dr.GetDouble(2).ToString("000000000"));
                            Automacao.setaValorNome("cdFilialAutorDestino", dr.GetDouble(3).ToString("0000"));
                            Automacao.setaValorNome("cdCtrlAutorDestino", dr.GetDouble(4).ToString("00"));
                            Automacao.ExecutarJS("obterPessoaJuridicaPorCNPJ('/gcpj')");
                        }
                        else
                        {
                            Automacao.selecionaCombo("cdCmbTipoEnvolvidoReu", "5", "value");
                            Automacao.espera();
                            Automacao.ExecutarJS("alterarTipoReu()");
                            Automacao.espera();
                            Automacao.setaValorNome("cdCodigoReuDestino", dr.GetDouble(2).ToString("000000000"));
                            Automacao.setaValorNome("cdFilialReuDestino", dr.GetDouble(3).ToString("0000"));
                            Automacao.setaValorNome("cdCtrlReuDestino", dr.GetDouble(4).ToString("00"));
                            Automacao.ExecutarJS("obterReuPessoaJuridicaPorCNPJ('/gcpj')");
                        }


                    }
                    Thread.Sleep(500);
                    if (Automacao.fecharPopUp("https://intranet85.net.bradesco.com.br/gcpj/servlet/procAlteracaoReuDeParaAction?acao=abrirPopUp"))
                    {
                        Conexao.executar("UPDATE TB_ENVOLVIDOS SET ATUALIZADO = TRUE, OBS = 'MAIS DE UM ENVOLVIDO LOCALIZADO COM O DOCUMENTO INFORMADO - VERIFIQUE' WHERE [CPF/CNPJ] = " + dr.GetDouble(2) + "and CNRO_PROCS_BDSCO = " + dr.GetDouble(0));
                    }
                    else if (Automacao.mataPossivelMsgBox())
                    {
                        //INSERE NO INFRA
                        Automacao.espera();
                        preencheGrade("Atualizando o envolvido: " + dr.GetString(1) + " incluindo o CPF/CNPJ já que ele não existe na base");
                        Automacao.navegar("https://intranet85.net.bradesco.com.br/gcpj/servlet/pessoaExternaBradescoAction?acao=carregar");
                        Automacao.espera();

                        Automacao.setaValorNome("cdPessoaExternaBradesco", idEnvolvido.Substring(0, 13));
                        Automacao.clicar("btoPesquisar");

                        if (Automacao.mataPossivelMsgBox())
                        {
                            Conexao.executar("UPDATE TB_ENVOLVIDOS SET ATUALIZADO = TRUE, OBS = 'NAO LOCALIZADO' WHERE [CPF/CNPJ] = " + dr.GetDouble(2) + "and CNRO_PROCS_BDSCO = " + dr.GetDouble(0));
                            continue;
                        }

                        Automacao.espera();
                        //Automacao.selecionarRadio(0, "selecionados");
                        if (Automacao.selecionarTabela("tab1", dr.GetString(1).Trim())==false)
                        {
                            Automacao.selecionarTabela("tab1", NomeEnvolvSemAcento );
                        }

                        Automacao.clicar("btoAlterar");
                        Automacao.espera();

                        if (dr.GetDouble(3).Equals(0))
                        {
                            Automacao.selecionaCombo("cdTipoPessoaExterna", "PESSOA FÍSICA");
                            Automacao.onChange("cdTipoPessoaExterna");
                            Automacao.setaValorNome("cdCpfPessoaExterna", dr.GetDouble(2).ToString("000000000"));
                            Automacao.setaValorNome("cdCtrlCpfPessoaExterna", dr.GetDouble(4).ToString("00"));
                        }
                        else
                        {
                            Automacao.selecionaCombo("cdTipoPessoaExterna", "PESSOA JURÍDICA");
                            Automacao.onChange("cdTipoPessoaExterna");
                            Automacao.setaValorNome("cdCnpjPessoaExterna", dr.GetDouble(2).ToString("00000000"));
                            Automacao.setaValorNome("cdFilialCnpjPessoaExterna", dr.GetDouble(3).ToString("0000"));
                            Automacao.setaValorNome("cdCtrlCnpjPessoaExterna", dr.GetDouble(4).ToString("00"));
                        }


                        String local = Automacao.retornaLocal();
                        Thread T1 = new Thread(new ThreadStart(Automacao.mataMsgBoxthread));
                        T1.Start();
                        Automacao.clicar("btoSalvar");
                        Automacao.mataPossivelMsgBox();
                        Automacao.espera();
                        T1.Interrupt();

                        if (local.Equals(Automacao.retornaLocal()))
                        {
                            Conexao.executar("UPDATE TB_ENVOLVIDOS SET ATUALIZADO = TRUE, OBS = 'DOCUMENTO INVALIDO' WHERE [CPF/CNPJ] = " + dr.GetDouble(2) + "and CNRO_PROCS_BDSCO = " + dr.GetDouble(0));
                        }
                        else
                        {
                            Conexao.executar("UPDATE TB_ENVOLVIDOS SET ATUALIZADO = TRUE, OBS = 'ATUALIZAÇÃO EFETUADA COM SUCESSO' WHERE [CPF/CNPJ] = " + dr.GetDouble(2) + "and CNRO_PROCS_BDSCO = " + dr.GetDouble(0));
                        }
                    }
                    else
                    {
                        //Verifica o nome se é igual
                        if (tipoEnvolvido == "A") 
                        { nomeEnvolvidLocalizado = Automacao.retornaValor("nmEnvolvidoAutorDestino"); }
                        else
                        { nomeEnvolvidLocalizado = Automacao.retornaValor("nmEnvolvidoReuDestino"); }

                        Automacao.clicar("btoSalvar");
                        Thread.Sleep(800);
                        String mensagem = Automacao.lerMsgBox();
                        if (mensagem.Length >= 33)
                        {
                            if (mensagem.Substring(0, 33).Equals("ENVOLVIDO JA FOI RELACIONADO COMO"))
                            {
                                Automacao.mataMsgBox();
                                //JA TEM CPF - CNPJ
                                Conexao.executar("UPDATE TB_ENVOLVIDOS SET ATUALIZADO = TRUE, OBS = 'DOCUMENTO PERTENCE A OUTRO ENVOLVIDO (" + nomeEnvolvidLocalizado  + ") JÁ RELACIONADO NESSE PROCESSO' WHERE [CPF/CNPJ] = " + dr.GetDouble(2) + "and CNRO_PROCS_BDSCO = " + dr.GetDouble(0));
                                preencheGrade("Documento pertence a outro envolvido já relacionado nesse processo.");
                                Automacao.espera();
                                continue;
                            }
                        }

                        if (dr.GetString(1).Trim().ToUpper() != nomeEnvolvidLocalizado.Trim().ToUpper() && NomeEnvolvSemAcento != nomeEnvolvidLocalizado.Trim().ToUpper())
                        {
                            Conexao.executar("UPDATE TB_ENVOLVIDOS SET ATUALIZADO = TRUE, OBS = 'ATUALIZAÇÃO EFETUADA COM SUCESSO. PORÉM COM UMA DIVERGÊNCIA NO NOME: " + nomeEnvolvidLocalizado + "' WHERE [CPF/CNPJ] = " + dr.GetDouble(2) + "and CNRO_PROCS_BDSCO = " + dr.GetDouble(0));
                        }
                        else
                        {
                            Conexao.executar("UPDATE TB_ENVOLVIDOS SET ATUALIZADO = TRUE, OBS = 'ATUALIZAÇÃO EFETUADA COM SUCESSO' WHERE [CPF/CNPJ] = " + dr.GetDouble(2) + "and CNRO_PROCS_BDSCO = " + dr.GetDouble(0));
                        }

                        Automacao.mataMsgBox();
                        continue;
                    }
                }
                catch (Exception)
                {
                    Conexao.executar("UPDATE TB_ENVOLVIDOS SET ATUALIZADO = TRUE, OBS = 'ERRO AO ATUALIZAR DOCUMENTO DO ENVOLVIDO' WHERE [CPF/CNPJ] = " + dr.GetDouble(2) + "and CNRO_PROCS_BDSCO = " + dr.GetDouble(0));
                    continue;
                }

            }
            carregaGridDados();
            gravaLOG("ATUALIZAÇÃO CPF/CNPJ ENVOLVIDO CONCLUÍDA - MÁQUINA " + qtdeMaquina);
            MessageBox.Show("Dados atualizados com sucesso!", "Pronto!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
        }

        public string ObterStringSemAcentosECaracteresEspeciais(string str)
        {
            /** Troca os caracteres acentuados por não acentuados **/
            string[] acentos = new string[] { "ç", "Ç", "á", "é", "í", "ó", "ú", "ý", "Á", "É", "Í", "Ó", "Ú", "Ý", "à", "è", "ì", "ò", "ù", "À", "È", "Ì", "Ò", "Ù", "ã", "õ", "ñ", "ä", "ë", "ï", "ö", "ü", "ÿ", "Ä", "Ë", "Ï", "Ö", "Ü", "Ã", "Õ", "Ñ", "â", "ê", "î", "ô", "û", "Â", "Ê", "Î", "Ô", "Û" };
            string[] semAcento = new string[] { "c", "C", "a", "e", "i", "o", "u", "y", "A", "E", "I", "O", "U", "Y", "a", "e", "i", "o", "u", "A", "E", "I", "O", "U", "a", "o", "n", "a", "e", "i", "o", "u", "y", "A", "E", "I", "O", "U", "A", "O", "N", "a", "e", "i", "o", "u", "A", "E", "I", "O", "U" };

            for (int i = 0; i < acentos.Length; i++)
            {
                str = str.Replace(acentos[i], semAcento[i]);
            }
            str = str.Trim();
            str = str.Replace("  ", " ");
            return str;
        }

        //INSERIR DEMAIS AUTORES
        private void btn_InsereEnvolvido_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Confirma a ação: INSERIR DEMAIS AUTORES contidos nesse arquivo importado?", "Atenção", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1) != DialogResult.Yes)
            {
                return;
            }
            if (IniciaProcessamento() == false) { return; }

            int qtdeMaquina = int.Parse(txtMaquina.Text);
            gravaLOG("INICIOU INCLUSÃO DEMAIS AUTORES - MÁQUINA " + qtdeMaquina);
            
            string mensagem;
            string NomeEnvolvSemAcento;
            String nomeEnvolvidLocalizado;
            int qtdOperacoesS5 = 0; // [S5] contador de registros processados nesta execução
            OleDbDataReader dr;
            dr = Conexao.consultar("SELECT * FROM TB_ENVOLVIDOS WHERE ATUALIZADO = FALSE AND MAQUINA = " + qtdeMaquina);

            while (dr.Read())
            {
                // [S5] Limitador de operações: para o processamento ao atingir o limite (ver LimitadorS5.cs)
                if (LimitadorS5.LimiteAtingido(qtdOperacoesS5))
                {
                    preencheGrade(LimitadorS5.MensagemLog());
                    gravaLOG("CONTROLE S5 - LIMITE DE OPERACOES ATINGIDO - INSERIR DEMAIS AUTORES - MÁQUINA " + qtdeMaquina);
                    break;
                }
                qtdOperacoesS5++;

                if (DateTime.Now.Hour.ToString() == "23")
                {
                    //vamos aguardar o outro dia
                    while (DateTime.Now.Hour.ToString() != "06")
                    {
                        //vamos aguardar ate as 06 da manha.
                        Thread.Sleep(3000000);
                    }
                }

                try
                {
                    NomeEnvolvSemAcento = ObterStringSemAcentosECaracteresEspeciais(dr.GetString(1));

                    //TELA - PROCESSO
                    Thread.Sleep(1000);
                    preencheGrade("Inserindo Autor " + dr.GetString(1));
                    Automacao.navegar("https://intranet85.net.bradesco.com.br/gcpj/servlet/cadMantemCadProcAlteracaoAction?acao=carregar&sessao=limpar&estimulo=alteracaoAtivos");
                    Automacao.espera();
                    Automacao.setaValorNome("nrProcessoBradesco", dr.GetDouble(0).ToString());
                    Automacao.espera();
                    Automacao.clicar("btoPesquisar");
                    if (Automacao.mataPossivelMsgBox())
                    {
                        Conexao.executar("UPDATE TB_ENVOLVIDOS SET ATUALIZADO = TRUE, OBS = 'FALHA AO PESQUISAR O PROCESSO. VERIFICAR.' WHERE [CPF/CNPJ] = " + dr.GetDouble(2) + "and CNRO_PROCS_BDSCO = " + dr.GetDouble(0));
                        continue;
                    }


                    Automacao.espera();
                    Automacao.ExecutarJS("paginar(2)");//Aba autor
                                                       //Automacao.espera();
                    Thread.Sleep(1000);
                    nomeEnvolvidLocalizado = "";
                    mensagem = "";

                    if (dr.GetDouble(3).Equals(0))
                    {
                        //select3
                        Automacao.selecionaCombo("select3", "PESSOA FISICA");
                        Automacao.espera();
                        Automacao.ExecutarJS("alteraDivTipo( 3, 5, true );");
                        Automacao.espera();
                        Automacao.setaValorNome("cdCodigo", dr.GetDouble(2).ToString("000000000"));
                        Automacao.espera();
                        Automacao.setaValorNome("cdCtrl", dr.GetDouble(4).ToString("00"));
                        Automacao.espera();
                        Automacao.clicar("btoCodigo");
                        Thread.Sleep(800);
                    }
                    else
                    {
                        //PJ
                        Automacao.selecionaCombo("select3", "PESSOA JURIDICA");
                        Automacao.espera();
                        Automacao.ExecutarJS("alteraDivTipo( 4, 5, true );");
                        Automacao.espera();
                        Automacao.setaValorNome("cdCodigo", dr.GetDouble(2).ToString("00000000"));
                        Automacao.espera();
                        Automacao.setaValorNome("cdFilial", dr.GetDouble(3).ToString("0000"));
                        Automacao.espera();
                        Automacao.setaValorNome("cdCtrl", dr.GetDouble(4).ToString("00"));
                        Automacao.espera();
                        Automacao.clicar("btoCodigo");
                        Thread.Sleep(800);
                    }

                    //Não localizou com o documento - INSERE NO INFRA
                    if (Automacao.mataPossivelMsgBox())
                    {
                        //cadastrarPF
                        if (dr.GetDouble(3).Equals(0))
                        {
                            Automacao.ExecutarJS("openTelasCadastro(5, 'PESSOA FISICA' );");
                            Automacao.espera();
                            Thread.Sleep(800);
                            mensagem = Automacao.cadastrarPF(NomeEnvolvSemAcento, dr.GetDouble(2), dr.GetDouble(4));
                        }
                        else //PJ
                        {
                            Automacao.ExecutarJS("openTelasCadastro(5, 'PESSOA JURIDICA' );");
                            Automacao.espera();
                            Thread.Sleep(800);
                            mensagem = Automacao.cadastrarPJ(NomeEnvolvSemAcento, dr.GetDouble(2), dr.GetDouble(3), dr.GetDouble(4));
                        }

                        if (mensagem == "NÃO FOI POSSÍVEL CADASTRAR ENVOLVIDO" || mensagem == "NÃO FOI POSSÍVEL CADASTRAR ENVOLVIDO - CPF / CNPJ INVALIDO")
                        {
                            Conexao.executar("UPDATE TB_ENVOLVIDOS SET ATUALIZADO = TRUE, OBS = '" + mensagem.ToUpper() + "' WHERE [CPF/CNPJ] = " + dr.GetDouble(2) + "and CNRO_PROCS_BDSCO = " + dr.GetDouble(0));
                            preencheGrade("Não foi possível cadastrar o envolvido " + dr.GetString(1) + " no INFRA.");
                            Automacao.espera();
                            continue;
                        }
                        nomeEnvolvidLocalizado = NomeEnvolvSemAcento;

                        //Automacao.espera();
                        //preencheGrade("Atualizando o envolvido: " + dr.GetString(1) + " incluindo o CPF/CNPJ já que ele não existe na base");
                        //Automacao.navegar("https://intranet85.net.bradesco.com.br/gcpj/servlet/pessoaExternaBradescoAction?acao=carregar");
                        //Automacao.espera();

                        //Automacao.clicar("btoIncluir");
                        //Automacao.espera();

                        //if (dr.GetDouble(3).Equals(0)) { Automacao.selecionaCombo("cdTipoPessoaExterna", "5", "value");} //PF
                        //else {Automacao.selecionaCombo("cdTipoPessoaExterna", "6", "value");}//PJ

                        //Automacao.espera();
                        //Automacao.ExecutarJS("tratarTipoPessoa();");
                        //Automacao.espera();

                        //Automacao.setaValorNome("nmPessoaExternaProcesso", dr.GetString(1).Trim());

                        //if (dr.GetDouble(3).Equals(0))
                        //{
                        //    Automacao.setaValorNome("cdCpfPessoaExterna", dr.GetDouble(2).ToString("000000000"));
                        //    Automacao.setaValorNome("cdControleCpfPessoaExterna", dr.GetDouble(4).ToString("00"));
                        //}
                        //else
                        //{
                        //    Automacao.setaValorNome("cdCnpjPessoaExterna", dr.GetDouble(2).ToString("00000000"));
                        //    Automacao.setaValorNome("cdFilialCnpjPessoaExterna", dr.GetDouble(3).ToString("0000"));
                        //    Automacao.setaValorNome("cdControleCnpjPessoaExterna", dr.GetDouble(4).ToString("00"));
                        //}

                        //Automacao.espera();
                        //Automacao.clicar("btoSalvar");
                        //mensagem = Automacao.lerMsgBox();
                        //if (mensagem.Substring(0, 13).Equals("Operação real"))
                        //{
                        //    //tudo  ok
                        //    preencheGrade("Envolvido " + dr.GetString(1) + " cadastrado com sucesso!");
                        //    Automacao.mataMsgBox();
                        //    goto retorno;
                        //}
                        //else if (mensagem.Substring(0, 11).Equals("CNPJ/CPF JA"))
                        //{
                        //    Automacao.mataMsgBox();
                        //    //JA TEM CPF - CNPJ
                        //    Conexao.executar("UPDATE TB_ENVOLVIDOS SET ATUALIZADO = TRUE, OBS = 'DOCUMENTO CADASTRADO PARA OUTRO ENVOLVIDO " + mensagem + "' WHERE [CPF/CNPJ] = " + dr.GetDouble(2) + "and CNRO_PROCS_BDSCO = " + dr.GetDouble(0));
                        //    preencheGrade("Envolvido " + dr.GetString(1) + " não foi cadastrado!");
                        //}
                        //else if (mensagem.Substring(0, 17).Equals("CPF/CNPJ INVALIDO"))
                        //{
                        //    Automacao.mataMsgBox();
                        //    //DOC INVALIDO - LIMPANDO OS CAMPOS
                        //    if (dr.GetDouble(3).Equals(0))
                        //    {
                        //        Automacao.setaValorNome("cdCpfPessoaExterna", "");
                        //        Automacao.setaValorNome("cdControleCpfPessoaExterna", "");
                        //    }
                        //    else
                        //    {
                        //        Automacao.setaValorNome("cdCnpjPessoaExterna", "");
                        //        Automacao.setaValorNome("cdFilialCnpjPessoaExterna", "");
                        //        Automacao.setaValorNome("cdControleCnpjPessoaExterna", "");
                        //    }
                        //    Automacao.clicar("btoSalvar");
                        //    Conexao.executar("UPDATE TB_ENVOLVIDOS SET ATUALIZADO = TRUE, OBS = 'CPF/CNPJ INVÁLIDO - ENVOLVIDO CADASTRADO SEM DOCUMENTO.' WHERE [CPF/CNPJ] = " + dr.GetDouble(2) + "and CNRO_PROCS_BDSCO = " + dr.GetDouble(0));
                        //    Automacao.mataMsgBox();
                        //    preencheGrade("Envolvido " + dr.GetString(1) + " cadastrado sem o documento!");
                        //    goto retorno;
                        //}
                        //else
                        //{
                        //    Conexao.executar("UPDATE TB_ENVOLVIDOS SET ATUALIZADO = TRUE, OBS = 'ERRO AO CADASTRAR ENVOLVIDO' WHERE [CPF/CNPJ] = " + dr.GetDouble(2) + "and CNRO_PROCS_BDSCO = " + dr.GetDouble(0));
                        //}

                    }
                    else //Localizou documento, pega o nome para verificar se é o mesmo 
                    {
                        nomeEnvolvidLocalizado = Automacao.retornaValor("dsNome");
                    }

                    Automacao.espera();
                    Thread.Sleep(500);
                    //VERIFICA SE ABRIU JANELA DE ERRO
                    if (Automacao.localizarJanela("ERRO!") == true)
                    {
                        //SE GEROU A JANELA ENTÃO GRAVAR COMO ERRO NA TABELA E PROSSEGUIR PARA O PROXIMO REGISTRO
                        Conexao.executar("UPDATE TB_ENVOLVIDOS SET ATUALIZADO = TRUE, OBS = 'ERRO TELA DO ENVOLVIDO' WHERE [CPF/CNPJ] = " + dr.GetDouble(2) + "and CNRO_PROCS_BDSCO = " + dr.GetDouble(0));
                        Automacao.FechaJanelaErro("ERRO!");
                        Automacao.localizarJanela("GCPJ");
                        goto proximo;
                    }

                    Automacao.clicar("btoIncluir");
                    String mensagem2 = Automacao.lerMsgBox();
                    if (mensagem2.Length >= 13)
                    {
                        if (mensagem2.Substring(0, 13).Equals("O envolvido j"))
                        {
                            Automacao.mataMsgBox();
                            //JA TEM CPF - CNPJ
                            Conexao.executar("UPDATE TB_ENVOLVIDOS SET ATUALIZADO = TRUE, OBS = 'O ENVOLVIDO JÁ ESTÁ RELACIONADO COMO AUTOR NESSE PROCESSO' WHERE [CPF/CNPJ] = " + dr.GetDouble(2) + "and CNRO_PROCS_BDSCO = " + dr.GetDouble(0));
                            preencheGrade("Envolvido " + dr.GetString(1) + " já está relacionado como autor nesse processo.");
                            Automacao.espera();
                            continue;
                        }
                    }

                    Automacao.espera();


                    String local = Automacao.retornaLocal();
                    Thread T1 = new Thread(new ThreadStart(Automacao.mataMsgBoxthread));
                    T1.Start();
                    Automacao.clicar("btoSalvar");
                    Automacao.mataPossivelMsgBox();
                    Automacao.espera();
                    T1.Interrupt();

                    if (local.Equals(Automacao.retornaLocal()))
                    {
                        Conexao.executar("UPDATE TB_ENVOLVIDOS SET ATUALIZADO = TRUE, OBS = 'ERRO AO SALVAR' WHERE [CPF/CNPJ] = " + dr.GetDouble(2) + "and CNRO_PROCS_BDSCO = " + dr.GetDouble(0));
                    }
                    else
                    {
                        if (dr.GetString(1).Trim().ToUpper() != nomeEnvolvidLocalizado.Trim().ToUpper() && NomeEnvolvSemAcento != nomeEnvolvidLocalizado.Trim().ToUpper())
                        {
                            Conexao.executar("UPDATE TB_ENVOLVIDOS SET ATUALIZADO = TRUE, OBS = 'ATUALIZAÇÃO EFETUADA COM SUCESSO. PORÉM COM UMA DIVERGÊNCIA NO NOME: " + nomeEnvolvidLocalizado + "' WHERE [CPF/CNPJ] = " + dr.GetDouble(2) + "and CNRO_PROCS_BDSCO = " + dr.GetDouble(0));
                        }
                        else
                        {
                            mensagem = mensagem.Replace("Operação realizada com sucesso.", "- ENVOLVIDO CADASTRADO NO INFRA - ");
                            Conexao.executar("UPDATE TB_ENVOLVIDOS SET ATUALIZADO = TRUE, OBS = 'ATUALIZAÇÃO EFETUADA COM SUCESSO " + mensagem + "' WHERE [CPF/CNPJ] = " + dr.GetDouble(2) + "and CNRO_PROCS_BDSCO = " + dr.GetDouble(0));
                        }
                    }

                proximo:
                    Automacao.espera();
                }
                catch (Exception)
                {
                    Conexao.executar("UPDATE TB_ENVOLVIDOS SET ATUALIZADO = TRUE, OBS = 'ERRO AO INSERIR ENVOLVIDO' WHERE [CPF/CNPJ] = " + dr.GetDouble(2) + "and CNRO_PROCS_BDSCO = " + dr.GetDouble(0));
                    continue;
                }  

            }

            preencheGrade("Processos atualizados com sucesso!");
            carregaGridDados();
            gravaLOG("INCLUSÃO DEMAIS AUTORES CONCLUÍDA  - MÁQUINA " + qtdeMaquina);
            MessageBox.Show("Dados atualizados com sucesso!", "Pronto!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);


        }

        //INSERIR PJ RÉU
        private void btn_InserePjReu_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Confirma a ação: INSERIR PJ RÉU contidos nesse arquivo importado?", "Atenção", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1) != DialogResult.Yes)
            {
                return;
            }
            if (IniciaProcessamento() == false) { return; }
            
            int qtdeMaquina = int.Parse(txtMaquina.Text);
            gravaLOG("INICIOU INCLUSÃO PJ RÉU - MÁQUINA " + qtdeMaquina);
           
            string mensagem;
            string NomeEnvolvSemAcento;
            String nomeEnvolvidLocalizado;
            int qtdOperacoesS5 = 0; // [S5] contador de registros processados nesta execução
            OleDbDataReader dr;
            dr = Conexao.consultar("SELECT * FROM TB_ENVOLVIDOS WHERE ATUALIZADO = FALSE AND MAQUINA = " + qtdeMaquina);

            while (dr.Read())
            {
                // [S5] Limitador de operações: para o processamento ao atingir o limite (ver LimitadorS5.cs)
                if (LimitadorS5.LimiteAtingido(qtdOperacoesS5))
                {
                    preencheGrade(LimitadorS5.MensagemLog());
                    gravaLOG("CONTROLE S5 - LIMITE DE OPERACOES ATINGIDO - INSERIR PJ REU - MÁQUINA " + qtdeMaquina);
                    break;
                }
                qtdOperacoesS5++;

                if (DateTime.Now.Hour.ToString() == "23")
                {
                    //vamos aguardar o outro dia
                    while (DateTime.Now.Hour.ToString() != "06")
                    {
                        //vamos aguardar ate as 06 da manha.
                        Thread.Sleep(3000000);
                    }
                }

                try
                {
                    NomeEnvolvSemAcento = ObterStringSemAcentosECaracteresEspeciais(dr.GetString(1));

                    //TELA - PROCESSO
                    Thread.Sleep(1000);
                    preencheGrade("Inserindo Réu " + dr.GetString(1));
                    Automacao.navegar("https://intranet85.net.bradesco.com.br/gcpj/servlet/cadMantemCadProcAlteracaoAction?acao=carregar&sessao=limpar&estimulo=alteracaoAtivos");
                    Automacao.espera();
                    Automacao.setaValorNome("nrProcessoBradesco", dr.GetDouble(0).ToString());
                    Automacao.espera();
                    Automacao.clicar("btoPesquisar");
                    if (Automacao.mataPossivelMsgBox())
                    {
                        Conexao.executar("UPDATE TB_ENVOLVIDOS SET ATUALIZADO = TRUE, OBS = 'FALHA AO PESQUISAR O PROCESSO. VERIFICAR.' WHERE [CPF/CNPJ] = " + dr.GetDouble(2) + "and CNRO_PROCS_BDSCO = " + dr.GetDouble(0));
                        continue;
                    }


                    Automacao.espera();
                    Automacao.ExecutarJS("paginar(3)");//Aba Réu
                                                       //Automacao.espera();
                    Thread.Sleep(1000);
                    nomeEnvolvidLocalizado = "";
                    mensagem = "";


                    if (dr.GetDouble(3).Equals(0))
                    {
                        //select3
                        Automacao.selecionaCombo("tiporeu", "PESSOA FISICA");
                        Automacao.espera();
                        Automacao.ExecutarJS("alteraDivTipo( 3, 5, true );");
                        Automacao.espera();
                        Automacao.setaValorNome("cdCodigo", dr.GetDouble(2).ToString("000000000"));
                        Automacao.espera();
                        Automacao.setaValorNome("cdCtrl", dr.GetDouble(4).ToString("00"));
                        Automacao.espera();
                        Automacao.clicar("btoCodigo");
                        Thread.Sleep(800);
                    }
                    else
                    {
                        //PJ
                        Automacao.selecionaCombo("tiporeu", "PESSOA JURIDICA");
                        Automacao.espera();
                        Automacao.ExecutarJS("alteraDivTipo( 4, 5, true );");
                        Automacao.espera();
                        Automacao.setaValorNome("cdCodigo", dr.GetDouble(2).ToString("00000000"));
                        Automacao.espera();
                        Automacao.setaValorNome("cdFilial", dr.GetDouble(3).ToString("0000"));
                        Automacao.espera();
                        Automacao.setaValorNome("cdCtrl", dr.GetDouble(4).ToString("00"));
                        Automacao.espera();
                        Automacao.clicar("btoCodigo");
                        Thread.Sleep(800);
                    }

                    //Não localizou com o documento - INSERE NO INFRA
                    if (Automacao.mataPossivelMsgBox())
                    {
                        //cadastrarPF
                        if (dr.GetDouble(3).Equals(0))
                        {
                            Automacao.ExecutarJS("openTelasCadastro(5, 'PESSOA FISICA' );");
                            Automacao.espera();
                            Thread.Sleep(800);
                            mensagem = Automacao.cadastrarPF(NomeEnvolvSemAcento, dr.GetDouble(2), dr.GetDouble(4));
                        }
                        else //PJ
                        {
                            Automacao.ExecutarJS("openTelasCadastro(5, 'PESSOA JURIDICA' );");
                            Automacao.espera();
                            Thread.Sleep(800);
                            mensagem = Automacao.cadastrarPJ(NomeEnvolvSemAcento, dr.GetDouble(2), dr.GetDouble(3), dr.GetDouble(4));
                        }

                        if (mensagem == "NÃO FOI POSSÍVEL CADASTRAR ENVOLVIDO" || mensagem == "NÃO FOI POSSÍVEL CADASTRAR ENVOLVIDO - CPF/CNPJ INVALIDO")
                        {
                            Conexao.executar("UPDATE TB_ENVOLVIDOS SET ATUALIZADO = TRUE, OBS = '" + mensagem.ToUpper() + "' WHERE [CPF/CNPJ] = " + dr.GetDouble(2) + "and CNRO_PROCS_BDSCO = " + dr.GetDouble(0));
                            preencheGrade("Não foi possível cadastrar o réu " + dr.GetString(1) + " no INFRA.");
                            Automacao.espera();
                            continue;
                        }
                        nomeEnvolvidLocalizado = NomeEnvolvSemAcento;

                    }
                    else //Localizou documento, pega o nome para verificar se é o mesmo 
                    {
                        nomeEnvolvidLocalizado = Automacao.retornaValor("dsNome");
                    }

                    Automacao.espera();
                    Thread.Sleep(500);
                    //VERIFICA SE ABRIU JANELA DE ERRO
                    if (Automacao.localizarJanela("ERRO!") == true)
                    {
                        //SE GEROU A JANELA ENTÃO GRAVAR COMO ERRO NA TABELA E PROSSEGUIR PARA O PROXIMO REGISTRO
                        Conexao.executar("UPDATE TB_ENVOLVIDOS SET ATUALIZADO = TRUE, OBS = 'ERRO TELA DO ENVOLVIDO' WHERE [CPF/CNPJ] = " + dr.GetDouble(2) + "and CNRO_PROCS_BDSCO = " + dr.GetDouble(0));
                        Automacao.FechaJanelaErro("ERRO!");
                        Automacao.localizarJanela("GCPJ");
                        goto proximo;
                    }

                    Automacao.clicar("btoIncluir");
                    String mensagem2 = Automacao.lerMsgBox();
                    if (mensagem2.Length >= 13)
                    {
                        if (mensagem2.Substring(0, 13).Equals("O envolvido j"))
                        {
                            Automacao.mataMsgBox();
                            //JA TEM CPF - CNPJ
                            Conexao.executar("UPDATE TB_ENVOLVIDOS SET ATUALIZADO = TRUE, OBS = 'O ENVOLVIDO JÁ ESTÁ RELACIONADO COMO RÉU NESSE PROCESSO' WHERE [CPF/CNPJ] = " + dr.GetDouble(2) + "and CNRO_PROCS_BDSCO = " + dr.GetDouble(0));
                            preencheGrade("Envolvido " + dr.GetString(1) + " já está relacionado como réu nesse processo.");
                            Automacao.espera();
                            continue;
                        }
                    }

                    Automacao.espera();


                    String local = Automacao.retornaLocal();
                    Thread T1 = new Thread(new ThreadStart(Automacao.mataMsgBoxthread));
                    T1.Start();
                    Automacao.clicar("btoSalvar");
                    Automacao.mataPossivelMsgBox();
                    Automacao.espera();
                    T1.Interrupt();

                    if (local.Equals(Automacao.retornaLocal()))
                    {
                        Conexao.executar("UPDATE TB_ENVOLVIDOS SET ATUALIZADO = TRUE, OBS = 'ERRO AO SALVAR' WHERE [CPF/CNPJ] = " + dr.GetDouble(2) + "and CNRO_PROCS_BDSCO = " + dr.GetDouble(0));
                    }
                    else
                    {
                        if (dr.GetString(1).Trim().ToUpper() != nomeEnvolvidLocalizado.Trim().ToUpper() && NomeEnvolvSemAcento != nomeEnvolvidLocalizado.Trim().ToUpper())
                        {
                            Conexao.executar("UPDATE TB_ENVOLVIDOS SET ATUALIZADO = TRUE, OBS = 'ATUALIZAÇÃO EFETUADA COM SUCESSO. PORÉM COM UMA DIVERGÊNCIA NO NOME: " + nomeEnvolvidLocalizado + "' WHERE [CPF/CNPJ] = " + dr.GetDouble(2) + "and CNRO_PROCS_BDSCO = " + dr.GetDouble(0));
                        }
                        else
                        {
                            mensagem = mensagem.Replace("Operação realizada com sucesso.", "- ENVOLVIDO CADASTRADO NO INFRA - ");
                            Conexao.executar("UPDATE TB_ENVOLVIDOS SET ATUALIZADO = TRUE, OBS = 'ATUALIZAÇÃO EFETUADA COM SUCESSO " + mensagem + "' WHERE [CPF/CNPJ] = " + dr.GetDouble(2) + "and CNRO_PROCS_BDSCO = " + dr.GetDouble(0));
                        }
                    }

                proximo:
                    Automacao.espera();
                }
                catch (Exception)
                {
                    Conexao.executar("UPDATE TB_ENVOLVIDOS SET ATUALIZADO = TRUE, OBS = 'ERRO AO INSERIR PJ RÉU' WHERE [CPF/CNPJ] = " + dr.GetDouble(2) + "and CNRO_PROCS_BDSCO = " + dr.GetDouble(0));
                    continue;
                }
               

            }

            preencheGrade("Processos atualizados com sucesso!");
            carregaGridDados();
            gravaLOG("INCLUSÃO PJ RÉU CONCLUÍDA - MÁQUINA " + qtdeMaquina);
            MessageBox.Show("Dados atualizados com sucesso!", "Pronto!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

        }
    }
}
