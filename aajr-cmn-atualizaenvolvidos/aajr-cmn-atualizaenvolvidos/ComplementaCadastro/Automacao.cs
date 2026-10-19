using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SHDocVw;
using System.Threading;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Windows.Automation;
using System.IO;
using ComplementaCadastro;



namespace ComplementaCadastro
{
    static class Automacao
    {

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = false)]
        private static extern IntPtr SendMessage(IntPtr hWnd, UInt32 Msg, int wParam, IntPtr lParam);

        [DllImport("user32.dll", EntryPoint = "FindWindow", SetLastError = true)]
        private static extern IntPtr FindWindowByCaption (IntPtr zero, String lpWindowName);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = false)] 
        private static extern IntPtr FindWindowEx (IntPtr parentHandle, IntPtr childAfter, String lclassName, String windowTitle);
        
        public const uint  WM_CLOSE = 0x0010;
        public const uint  BM_CLICK = 0x00F5;
        public const UInt32   WM_ACTIVATE = 0x0006;
        public const int  WA_ACTIVE = 1;

        public static AutomationElement aeDesktop;
        public static AutomationElement frmSelecionar;
        public static AutomationElement arquivo;
        public static AutomationElement btnAbrir;

        public static ValuePattern arquivoTexto;

        public static InvokePattern btnAbrirClick;

        public static bool telaCompleta {set; get;}
        static InternetExplorer instancia { set; get; }
        static InternetExplorer instAnexo { set; get; }
        static InternetExplorer instCad { set; get; }


        public static bool localizarJanela(String nome_janela)
        {
            foreach (InternetExplorer ie in new ShellWindows())
            {
                if (ie.LocationName.IndexOf(nome_janela) >= 0)
                {
                    instancia = ie;
                    instancia.Visible = true;
                    return true;
                }
            }
            return false;
        }

        public static bool novaJanela(String url)
        {
            instancia = new InternetExplorer();
            try
            {
                instancia.Navigate(url);
            }
            catch (Exception)
            {

                return false;
            }
            
            return true;
        }

        public static bool setaVisibilidade(Boolean visibilidade)
        {
            try
            {
                instancia.Visible = visibilidade;
            }
            catch (Exception)
            {
                return false;
            }
            return true;
        }

        public static bool navegar(String url)
        {
            try
            {
                if (telaCompleta)
                {
                    instancia.Document.frames(2).navigate2(url);
                }
                else
                {
                    instancia.Navigate2(url);
                }
                
            }
            catch (Exception)
            {
                return false;
            }
            return true;
        }

        public static bool setaValorNome(String nome_componente, String valor)
        {
            try
            {
                if (telaCompleta)
                {
                    instancia.Document.frames(2).Document.forms(0).elements(nome_componente).innertext = valor;
                }
                else
                {
                    instancia.Document.forms(0).elements(nome_componente).innertext = valor;
                }
                
            }
            catch (Exception)
            {
                return false;   
            }
            return true;
        }

        public static bool setaValorId(String id_componente, String valor)
        {
            try
            {
                if (telaCompleta)
                {
                    instancia.Document.frames(2).Document.getElementById(id_componente).innertext = valor;
                }
                else
                {
                    instancia.Document.getElementById(id_componente).innertext = valor;
                }
                
            }
            catch (Exception)
            {
                return false;
            }
            return true;
        }

        public static bool clicar(String botao)
        {
            try
            {
                if (telaCompleta)
                {
                    instancia.Document.frames(2).Document.forms(0).elements(botao).click();
                }
                else
                {
                    instancia.Document.forms(0).elements(botao).click();
                }
                
            }
            catch (Exception)
            {
                MessageBox.Show("erro!");
                return false;
            }
            return true;
        }

        public static bool onClick(String botao)
        {
            try
            {
                if (telaCompleta)
                {
                    instancia.Document.frames(2).Document.forms(0).elements(botao).onclick();
                }
                else
                {
                    instancia.Document.forms(0).elements(botao).onclick();
                }
                
            }
            catch (Exception)
            {

                return false;
            }
            return true;
        }

        public static bool espera()
        {
            //Thread.Sleep(500);
            while ((instancia.ReadyState == tagREADYSTATE.READYSTATE_INTERACTIVE) || (instancia.ReadyState == tagREADYSTATE.READYSTATE_LOADING) || (instancia.ReadyState == tagREADYSTATE.READYSTATE_UNINITIALIZED))
            {
                //aguarde
            }

            while (instancia.Busy)
            {
                //aguarde
            }
            return true;
        }

        public static bool esperaAnexo()
        {
            Thread.Sleep(500);
            while ((instAnexo.ReadyState == tagREADYSTATE.READYSTATE_INTERACTIVE) || (instAnexo.ReadyState == tagREADYSTATE.READYSTATE_LOADING) || (instAnexo.ReadyState == tagREADYSTATE.READYSTATE_UNINITIALIZED))
            {
                //aguarde
            }

            while (instAnexo.Busy)
            {
                //aguarde
            }
            return true;
        }
        public static bool fecharPopUp(string url)
        {
            foreach (InternetExplorer ie in new ShellWindows())
            {
                if (ie.LocationURL.Contains(url))
                {
                    ie.Quit();
                    return true;
                }
            }
            return false;
        }

        public static string lerMsgBox(bool cadastro = false)
        {
            IntPtr winHwnd = IntPtr.Zero;
            while (winHwnd == IntPtr.Zero)
            {
                winHwnd = FindWindowByCaption(IntPtr.Zero, "Mensagem da página da web");
                if (cadastro)
                {
                    if (instCad.Busy == false)
                    {
                        Thread.Sleep(1000);
                        return "";
                    }
                }
                else
                {
                    if (instancia.Busy == false )
                    {
                        Thread.Sleep(1000);
                        return "";
                    }
                }
               
            }
            Thread.Sleep(1000);

            AutomationElement janela;
            AutomationElementCollection textos;
            String retorno = "";

            janela = AutomationElement.FromHandle(winHwnd);
            textos = janela.FindAll(TreeScope.Children, new PropertyCondition(AutomationElement.ClassNameProperty, "Static"));

            foreach (AutomationElement item in textos)
            {
                if (item.Current.LocalizedControlType.Equals("text"))
                {
                    retorno += item.Current.Name.ToString().Replace("\n", "");
                }
            }

            return retorno;
        }

        public static bool mataMsgBox()
        {
            IntPtr winHwnd = IntPtr.Zero;
            while (winHwnd == IntPtr.Zero)
            {
                winHwnd = FindWindowByCaption(IntPtr.Zero, "Mensagem da página da web");
            }

            Thread.Sleep(1000);
            IntPtr btn = FindWindowEx(winHwnd, IntPtr.Zero, "Button", "OK");
            SendMessage(btn, WM_ACTIVATE, WA_ACTIVE, IntPtr.Zero);
            SendMessage(btn, BM_CLICK, 0, IntPtr.Zero);
            return true;
        }

        public static void mataMsgBoxthread()
        {
            IntPtr winHwnd = IntPtr.Zero;
            while (winHwnd == IntPtr.Zero)
            {
                winHwnd = FindWindowByCaption(IntPtr.Zero, "Mensagem da página da web");
            }

            Thread.Sleep(1000);
            IntPtr btn = FindWindowEx(winHwnd, IntPtr.Zero, "Button", "OK");
            SendMessage(btn, WM_ACTIVATE, WA_ACTIVE, IntPtr.Zero);
            SendMessage(btn, BM_CLICK, 0, IntPtr.Zero);
        }

        public static bool mataPossivelMsgBox()
        {
            Thread.Sleep(1000);
            IntPtr winHwnd = IntPtr.Zero;
            while (winHwnd == IntPtr.Zero)
            {
                winHwnd = FindWindowByCaption(IntPtr.Zero, "Mensagem da página da web");
                if (instancia.Busy == false)
                {
                    Thread.Sleep(1000);
                    return false;
                }
            }

            Thread.Sleep(1000);
            IntPtr btn = FindWindowEx(winHwnd, IntPtr.Zero, "Button", "OK");
            SendMessage(btn, WM_ACTIVATE, WA_ACTIVE, IntPtr.Zero);
            SendMessage(btn, BM_CLICK, 0, IntPtr.Zero);
            return true;
        }

        public static bool ExecutarJS(String comando)
        {
            if (telaCompleta)
            {
                instancia.Document.frames(2).Document.parentWindow.execScript(comando, "JavaScript");
            }
            else
            {
                //MessageBox.Show(instancia.Document.Url.ToString());
                instancia.Document.parentWindow.execScript(comando, "JavaScript");
            }
            return true;
        }

        public static String retornaValor(String nome)
        {
            if (telaCompleta)
            {
                return instancia.Document.frames(2).Document.forms(0).elements(nome).innertext();
            }
            else
            {
               // return instancia.Document.all.Item(nome).innertext();
                return instancia.Document.all.item(nome).value;
            }
            
        }

        public static String retornaValue(String nome)
        {
            if (telaCompleta)
            {
                return instancia.Document.frames(2).Document.forms(0).elements(nome).value();
            }
            else
            {
                return instancia.Document.forms(0).elements(nome).value();
            }
            
        }

        public static HtmlElement retornaObjeto(String tag, String classe)
        {
            
            HtmlElementCollection objs;
            objs = (HtmlElementCollection)instancia.Document.GetElementsByTagName(tag);

            foreach (Object obj in objs)
	        {
                MessageBox.Show(obj.ToString());
	        }
            return null;
	    }

        public static void onBlur(String elemento)
        {
            if (telaCompleta)
            {
                instancia.Document.frames(2).Document.forms(0).elements(elemento).onblur();
            }
            else
            {
                instancia.Document.forms(0).elements(elemento).onblur();
            }
            espera();
                
        }

        public static string retornaLocal()
        {
            return instancia.LocationURL;
        }

        public static void onChange(String elemento)
        {
            if (telaCompleta)
            {
                instancia.Document.frames(2).Document.forms(0).elements(elemento).onchange();
            }
            else
            {
                instancia.Document.forms(0).elements(elemento).onchange();
            }
            espera();

        }

        public static void selecionaCombo(String nome, String valor, String chave = "")
        {
            if (telaCompleta)
            {
                for (int x = 0; x <= 99; x++)
                {
                    if (chave == "")
                    {
                        if (instancia.Document.frames(2).Document.forms(0).elements(nome)(x).innertext().Trim() == valor.Trim())
                        {
                            instancia.Document.frames(2).Document.forms(0).elements(nome)(x).selected = true;
                            instancia.Document.frames(2).Document.forms(0).elements(nome).onblur();
                            espera();
                            instancia.Document.frames(2).Document.forms(0).elements(nome).onchange();
                            espera();
                            return;
                        }

                    }
                    else
                    {
                        if (instancia.Document.frames(2).Document.forms(0).elements(nome)(x).getAttribute(chave).Trim() == valor.Trim())
                        {
                            instancia.Document.frames(2).Document.forms(0).elements(nome)(x).selected = true;
                            instancia.Document.frames(2).Document.forms(0).elements(nome).onblur();
                            espera();
                            instancia.Document.frames(2).Document.forms(0).elements(nome).onchange();
                            espera();
                            return;
                        }
                    }
                }
            }
            else
            {
                for (int x = 0; x <= 99; x++)
                {
                    if (chave == "")
                    {
                        // RO 2632561 - KAYLANE ALVES MACHADO
                        // Inclusão do.tri porque a captura do innertext do site estava vindo com um espaço na frente, impedindo o fluxo de entrar no if
                        if (instancia.Document.forms(0).elements(nome)(x).innertext().Trim() == valor.Trim())
                        {
                            instancia.Document.forms(0).elements(nome)(x).selected = true;
                            instancia.Document.forms(0).elements(nome).onblur();
                            espera();
                            instancia.Document.forms(0).elements(nome).onchange();
                            espera();
                            return;
                        }

                    }
                    else
                    {
                        if (instancia.Document.forms(0).elements(nome)(x).getAttribute(chave).Trim() == valor.Trim())
                        {
                            instancia.Document.forms(0).elements(nome)(x).selected = true;
                            instancia.Document.forms(0).elements(nome).onblur();
                            espera();
                            instancia.Document.forms(0).elements(nome).onchange();
                            espera();
                            return;
                        }
                    }
                }
            }

            
        }

        public static void criarJanelaAnexo(String gcpj)
        {
            instAnexo = new InternetExplorer();
            instAnexo.Navigate("https://intranet85.net.bradesco.com.br/gcpj/servlet/cadProcsAnexosAction?acao=carregar&cdNumeroProcessoBradesco=" + gcpj + "&permitirSalvar=1&tituloPagina=Automação Captura de Iniciais");
            instAnexo.Visible = true;
            Thread.Sleep(2000);
        }

        public static bool inserirAnexo(String nomeDocumento, int tipoAnexo, string Diretorio, string CaminhoArquivo, String gcpj)
        {
            instAnexo = new InternetExplorer();
            instAnexo.Navigate("https://intranet85.net.bradesco.com.br/gcpj/servlet/cadProcsAnexosAction?acao=carregar&cdNumeroProcessoBradesco=" + gcpj + "&permitirSalvar=1&tituloPagina=Automação Captura de Iniciais");

            Thread.Sleep(2000);

            foreach (InternetExplorer ie in new ShellWindows())
            {
                if (ie.LocationURL.IndexOf("https://intranet85.net.bradesco.com.br/gcpj/servlet/cadProcsAnexosAction?acao=carregar&cdNumeroProcessoBradesco=" + gcpj) >= 0)
                {
                    instAnexo = ie;
                    break;
                }
            }
            esperaAnexo();

            Thread tr2 = new Thread(clicarAnexar);
        denovo:
            instAnexo.Document.forms(0).elements("nmAnexoProcesso").innertext = nomeDocumento;
            for (int x = 1; x <= 999; x++)
            {
                if (instAnexo.Document.forms(0).elements("cmdAnexos")(x).getAttribute("cdTipoAnexo") == tipoAnexo.ToString())
                {
                    instAnexo.Document.forms(0).elements("cmdAnexos")(x).selected = true;
                    instAnexo.Document.forms(0).elements("cmdAnexos").onblur();
                    esperaAnexo();
                    instAnexo.Document.forms(0).elements("cmdAnexos").onchange();
                    esperaAnexo();
                    instAnexo.Document.forms(0).elements("cdSelecionadoAnexo").value = tipoAnexo.ToString();
                    instAnexo.Document.forms(0).elements("dsTipoAnexo").value = instAnexo.Document.forms(0).elements("cmdAnexos")(x).innertext();
                    break;
                }
            }

            //TRATAMENTO DO NOME DO ARQUIVO
            String extensao;
            try
            {
                extensao = "pdf";
                Double.Parse(CaminhoArquivo.Split('_')[0]);
                DirectoryInfo dirbase = new DirectoryInfo(Diretorio);
                FileInfo[] taskFiles = dirbase.GetFiles(CaminhoArquivo.Split('_')[0] + "*." + extensao);
                taskFiles[0].CopyTo(Diretorio + gcpj + "." + extensao);
            }
            catch (Exception)
            {
                try
                {
                    extensao = Path.GetExtension(Diretorio + CaminhoArquivo);
                    extensao.Replace(".","");
                    File.Copy(Diretorio + CaminhoArquivo, Diretorio + gcpj + "." + extensao);
                }
                catch (Exception)
                {
                    
                    //NAO LOCALIZEI A INICIAL
                    Conexao.executar("UPDATE SYS_IMPORTACAO SET PRE_CADASTRADO = TRUE, GCPJ = 9999999 WHERE IDENTIFICADOR_ANEXO = " + CaminhoArquivo.Split('_')[0]);
                    instAnexo.Quit();
                    return false;
                }
                
            }

            tr2.Start();
            Thread.Sleep(2000);
again:
            try
            {
                aeDesktop = AutomationElement.RootElement;
                frmSelecionar = aeDesktop.FindFirst(TreeScope.Descendants, new PropertyCondition(AutomationElement.NameProperty, "Escolher arquivo a carregar"));
                arquivo = frmSelecionar.FindFirst(TreeScope.Descendants, new PropertyCondition(AutomationElement.ClassNameProperty, "Edit"));

                arquivoTexto = (ValuePattern)arquivo.GetCurrentPattern(ValuePattern.Pattern);
                arquivoTexto.SetValue(Diretorio + gcpj + "." + extensao);

                btnAbrir = frmSelecionar.FindFirst(TreeScope.Children, new PropertyCondition(AutomationElement.NameProperty, "Abrir"));
                btnAbrirClick = (InvokePattern)btnAbrir.GetCurrentPattern(InvokePattern.Pattern);
                btnAbrirClick.Invoke();
            }
            catch (Exception)
            {
                goto again;
            }

            instAnexo.Document.forms(0).elements("btoIncluir").click();
            if (mataPossivelMsgBox())
            {
                goto denovo;
            }
            File.Delete(Diretorio + gcpj + ".pdf");
            //GARANTIR A ROTINA.... E DEPOIS DELETAR TUDO!
            //File.Delete(Diretorio + CaminhoArquivo);
           // instAnexo.Document.forms(0).elements("btoSalvar").click();
           // mataMsgBox();
            instAnexo.Quit();
            return true;
        }

        public static void clicarAnexar()
        {
            instAnexo.Document.forms(0).elements("formFile").click();
        }

        public static void fecharJanelaAnexo()
        {
            instAnexo.Quit();
        }

        public static bool FechaJanelaErro(String nome_janela)
        {
            foreach (InternetExplorer ie in new ShellWindows())
            {
                if (ie.LocationName.IndexOf(nome_janela) >= 0)
                {
                    instancia = ie;
                    instancia.Quit();
                    return true;
                }
            }
            return false;
        }

        public static string cadastrarPF(String nome, double cpf, double digito)
        {
            String mensagem = "";
            bool achou = false;

            foreach (InternetExplorer ie in new ShellWindows())
            {
                if (ie.LocationURL.IndexOf("https://intranet85.net.bradesco.com.br/gcpj/servlet/cadMantemCadProcTrabAdmReuAction?acao=novaPessoaExterna&cdIndicadorOrigemChamado=S&cdIndicativoPopUp=C&origem=origemCivel&dsTipoPessoa=PESSOA FISICA") >= 0)
                {
                    instCad = ie;
                    instCad.Visible = true;
                    achou = true;
                    break;
                }
            }

            //iniciando o cadastro
            if (achou)
            {
                instCad.Document.forms(0).elements("nmPessoaExternaProcesso").innertext = nome;
                if (cpf != 0)
                {
                    instCad.Document.forms(0).elements("cdCpfPessoaExterna").innertext = cpf;
                    instCad.Document.forms(0).elements("cdCtrlCpfPessoaExterna").innertext = digito;
                }

                instCad.Document.forms(0).elements("btoSalvar").click();
                Thread.Sleep(1000);
                mensagem = lerMsgBox(true);
                if (mensagem.Substring(0, 17).Equals("CPF/CNPJ INVALIDO"))
                {
                    mensagem = "NÃO FOI POSSÍVEL CADASTRAR ENVOLVIDO - CPF/CNPJ INVALIDO";
                    mataMsgBox();
                    instCad.Quit();

                    //mataMsgBox();
                    ////DOC INVALIDO - LIMPANDO OS CAMPOS
                    //instCad.Document.forms(0).elements("cdCpfPessoaExterna").innertext = "";
                    //instCad.Document.forms(0).elements("cdControleCpfPessoaExterna").innertext = "";

                    //instCad.Document.forms(0).elements("btoSalvar").click();

                    //mensagem = lerMsgBox(true) + " - CPF INVÁLIDO - ENVOLVIDO CADASTRADO SEM CPF";
                }
                else
                { mataMsgBox(); }

                return mensagem;
            }
            else
            {
                return "NÃO FOI POSSÍVEL CADASTRAR ENVOLVIDO";
            }
           
        }

        public static string cadastrarPJ(String nome, double cpf, double filial, double digito)
        {
            String mensagem = "";
            bool achou = false;

            foreach (InternetExplorer ie in new ShellWindows())
            {
                if (ie.LocationURL.IndexOf("https://intranet85.net.bradesco.com.br/gcpj/servlet/cadMantemCadProcTrabAdmReuAction?acao=novaPessoaExterna&cdIndicadorOrigemChamado=S&cdIndicativoPopUp=C&origem=origemCivel&dsTipoPessoa=PESSOA JURIDICA") >= 0)
                {                           
                    instCad = ie;
                    instCad.Visible = true;
                    achou = true;
                    break;
                }
            }

            //iniciando o cadastro
            if (achou)
            {
                instCad.Document.forms(0).elements("nmPessoaExternaProcesso").innertext = nome;

                if (cpf != 0)
                {
                    instCad.Document.forms(0).elements("cdCnpjPessoaExterna").innertext = cpf;
                    instCad.Document.forms(0).elements("cdFilialCnpjPessoaExterna").innertext = filial;
                    instCad.Document.forms(0).elements("cdCtrlCnpjPessoaExterna").innertext = digito;
                }
                instCad.Document.forms(0).elements("btoSalvar").click();
                Thread.Sleep(1000);
                mensagem = lerMsgBox(true);
                if (mensagem.Substring(0, 17).Equals("CPF/CNPJ INVALIDO"))
                {
                    mensagem = "NÃO FOI POSSÍVEL CADASTRAR ENVOLVIDO - CPF/CNPJ INVALIDO";
                    mataMsgBox();
                    instCad.Quit();
                }
                else
                { mataMsgBox(); }

                return mensagem;
            }
            else
            {
                return "NÃO FOI POSSÍVEL CADASTRAR ENVOLVIDO";
            }
        }

        public static void selecionarRadio(int sequencia, string componente)
        {
            try
            {
                instancia.Document.forms(0).elements(componente)(sequencia).SetAttribute("checked", "true");
            }
            catch (Exception)
            {
                try
                {
                    instancia.Document.forms(0).elements(componente).SetAttribute("checked", "true");
                }
                catch (Exception)
                {
                    
                }
                
            }

        }

        public static string selecionarTabelaEnvolvido(String classe, String pesquisa1, String pesquisa2)
        {
            try
            {
                string idEnvolvido = "";
                string cpf_cnpj = "";
                mshtml.HTMLDocument corpo = instancia.Document;
                mshtml.IHTMLElementCollection Tabelas;
                Tabelas = instancia.Document.getElementsByTagName("table");
                foreach (mshtml.HTMLTable Tabela in Tabelas)
                {
                    if (Tabela.className != null)
                    {
                        if (Tabela.className.ToString().Equals(classe))
                        {
                            foreach (mshtml.HTMLTableRow Linha in Tabela.rows)
                            {
                                foreach (mshtml.HTMLTableCell Celula in Linha.cells)
                                {
                                    if (Celula.innerText != null)
                                    {
                                        if (Celula.innerText.ToString().Trim().Equals(pesquisa1) || Celula.innerText.ToString().Trim().Equals(pesquisa2))
                                        {
                                            idEnvolvido = Celula.innerHTML;
                                            int pos = idEnvolvido.IndexOf("abrirPopUpDetalhes('00") + "abrirPopUpDetalhes('00".Length;
                                            // idEnvolvido = Celula.innerHTML.ToString().Replace("< A class=lnkTab1 href =\"javascript:abrirPopUpDetalhes('","");
                                            idEnvolvido = idEnvolvido.Substring(pos, 15);

                                            if (Linha.all.tags("td")(2).innertext != null)
                                            {
                                                cpf_cnpj = Linha.all.tags("td")(2).innertext;
                                                cpf_cnpj = cpf_cnpj.Replace(".", "").Replace("-", "").Replace("/", "");
                                                idEnvolvido = idEnvolvido + "-" + cpf_cnpj;
                                            }
                                            //Pega a inicial pra ver se é autor ou réu
                                            idEnvolvido = idEnvolvido + Linha.all.tags("td")(3).innertext.Substring(0, 1);

                                            mshtml.IHTMLElementCollection checkboxes = Linha.getElementsByTagName("input");
                                            foreach (mshtml.HTMLInputElement checkbox in checkboxes)
                                            {
                                                checkbox.setAttribute("checked", "true");
                                                return idEnvolvido;
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }

                }
                return "";
            }
            catch (Exception)
            {

                return "ERRO";
            }
        }

        public static bool selecionarTabela(String classe, String pesquisa)
        {
            mshtml.HTMLDocument corpo = instancia.Document;
            mshtml.IHTMLElementCollection Tabelas;
            Tabelas = instancia.Document.getElementsByTagName("table");
            foreach (mshtml.HTMLTable Tabela in Tabelas)
            {
                if (Tabela.className != null)
                {
                    if (Tabela.className.ToString().Equals(classe))
                    {
                        foreach (mshtml.HTMLTableRow Linha in Tabela.rows)
                        {
                            foreach (mshtml.HTMLTableCell Celula in Linha.cells)
                            {
                                if (Celula.innerText != null)
                                {
                                    if (Celula.innerText.ToString().Trim().Equals(pesquisa))
                                    {
                                        mshtml.IHTMLElementCollection checkboxes = Linha.getElementsByTagName("input");
                                        foreach (mshtml.HTMLInputElement checkbox in checkboxes)
                                        {
                                            checkbox.setAttribute("checked", "true");
                                            return true;
                                        }
                                    }
                                }
                            }
                        }
                    }
                }

            }
            return false;
        }

        public static bool selecionarTabela(String classe, String pesquisa1, String pesquisa2)
        {
            mshtml.HTMLDocument corpo = instancia.Document;
            mshtml.IHTMLElementCollection Tabelas;
            Tabelas = instancia.Document.getElementsByTagName("table");
            foreach (mshtml.HTMLTable Tabela in Tabelas)
            {
                if (Tabela.className != null)
                {
                    if (Tabela.className.ToString().Equals(classe))
                    {
                        foreach (mshtml.HTMLTableRow Linha in Tabela.rows)
                        {
                            bool localizado1 = false;
                            bool localizado2 = false;
                            foreach (mshtml.HTMLTableCell Celula in Linha.cells)
                            {
                                if (Celula.innerText != null)
                                {
                                    if (Celula.innerText.ToString().Trim().Equals(pesquisa1))
                                    {
                                        localizado1 = true;
                                    }
                                    if (Celula.innerText.ToString().Trim().Equals(pesquisa2))
                                    {
                                        localizado2 = true;
                                    }
                                    if (localizado1 == true && localizado2 == true)
                                    {
                                        mshtml.IHTMLElementCollection checkboxes = Linha.getElementsByTagName("input");
                                        foreach (mshtml.HTMLInputElement checkbox in checkboxes)
                                        {
                                            checkbox.setAttribute("checked", "true");
                                            return true;
                                        }
                                    }
                                }
                            }
                        }
                    }
                }

            }
            return false;
        }

        public static string RetornaPrimeiro(String classe)
        {
            mshtml.HTMLDocument corpo = instancia.Document;
            mshtml.IHTMLElementCollection Tabelas;
            Tabelas = instancia.Document.getElementsByTagName("table");
            foreach (mshtml.HTMLTable Tabela in Tabelas)
            {
                if (Tabela.className != null)
                {
                    if (Tabela.className.ToString().Equals(classe))
                    {
                        foreach (mshtml.HTMLTableRow Linha in Tabela.rows)
                        {
                            foreach (mshtml.HTMLTableCell Celula in Linha.cells)
                            {
                                if (Celula.innerText != null)
                                {
                                    try
                                    {
                                        decimal.Parse(Celula.innerText);
                                        return Celula.innerText;
                                    }
                                    catch (Exception)
                                    {
                                        //erro segue
                                    }
                                }
                            }
                        }
                    }
                }

            }
            return "";
        }


    }
}
