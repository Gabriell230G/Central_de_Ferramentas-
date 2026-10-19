using System.Windows.Forms;

namespace ComplementaCadastro
{
    // =========================================================================
    //  CONTROLE S5 - LIMITADOR DE OPERAÇÕES
    //  Normativo 05.1073 - Sistemas/Aplicativos Mantidos em Ambiente Departamental
    // -------------------------------------------------------------------------
    //  O QUE FAZ
    //    Limita quantos registros o robô pode processar no GCPJ em UMA execução.
    //    Quando o limite é atingido:
    //      1. o laço da automação para (break);
    //      2. o usuário recebe o aviso "Processamento Limitado - Controle S5";
    //      3. os registros não processados continuam com ATUALIZADO = FALSE na TB_ENVOLVIDOS e entram na próxima execução.
    //
    //  COMO MUDAR O LIMITE
    //    Altere SOMENTE a constante LIMITE_OPERACOES abaixo.
    //
    //  ONDE ELE É USADO
    //    Ctrl+Shift+F no Visual Studio e pesquise por "[S5]".
    //      - Form1.cs > btnAtualizaCPF_Click (D@KQ - Atualizar CPF/CNPJ Envolvido)
    //      - Form1.cs > btn_InsereEnvolvido_Click (D@KV - Inserir Demais Autores)
    //      - Form1.cs > btn_InserePjReu_Click (D@LM - Inserir PJ Réu)
    // =========================================================================
    static class LimitadorS5
    {
        /// <summary>Máximo de registros processados no GCPJ por execução.</summary>
        public const int LIMITE_OPERACOES = 5000;

        /// <summary>
        /// Retorna true quando a execução já chegou no limite do Controle S5
        /// e, nesse caso, avisa o usuário. Uso dentro do laço:
        ///     if (LimitadorS5.LimiteAtingido(qtdOperacoesS5)) break;
        ///     qtdOperacoesS5++;
        /// </summary>
        /// <param name="qtdOperacoes">Quantos registros já foram processados nesta execução.</param>
        public static bool LimiteAtingido(int qtdOperacoes)
        {
            if (qtdOperacoes < LIMITE_OPERACOES) return false;

            MessageBox.Show(
                "O limite de processamento de " + LIMITE_OPERACOES.ToString("N0") + " registros foi atingido.\n\n" +
                "Os itens restantes continuam pendentes e podem ser processados em uma nova execução.",
                "Processamento Limitado - Controle S5", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            return true;
        }

        /// <summary>Texto padrão para registrar na grade/log quando o limite é atingido.</summary>
        public static string MensagemLog()
        {
            return "Controle S5: limite de " + LIMITE_OPERACOES.ToString("N0") + " registros atingido. Processamento interrompido.";
        }
    }
}
