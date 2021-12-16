using System;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Text;
using System.Windows.Forms;

namespace EmailProject
{
    public partial class MainForm : Form
    {
        public MainForm()
        {
            InitializeComponent();
            CenterToScreen();
        }

        private void CbAnexos_CheckedChanged(object sender, EventArgs e)
        {
            //verificando se o check box está selecionado
            //se não estiver, ele bloqueia o acesso a função de carregar um anexo no e-mail
            gbAnexo.Enabled = (sender as CheckBox).Checked;
        }

        private void BtnCarregarArquivo_Click(object sender, EventArgs e)
        {
            //intânciando uma nova caixa de diálogo para carregar arquivos
            using (var openDialog = new OpenFileDialog())
            {
                //definindo o título da caixa de diálogo
                openDialog.Title = "Anexar Arquivo(s)";

                //definindo que a caixa de diálogo poderá carregar mais de 1 arquivo
                openDialog.Multiselect = true;

                //verificando se um arquivo foi carregado
                if (openDialog.ShowDialog() == DialogResult.OK)
                {
                    //percorrendo os arquivos que foram carregados pelo usuário
                    for (int i = 0; i < openDialog.FileNames.Count(); i++)
                    {
                        //if que erá executado somente no primeiro índice
                        if (i == 0)
                        {
                            tbAnexos.Text = openDialog.FileNames[i];
                            continue;
                        }

                        //adicionando um ";" como um separador entre cada arquivo carregado
                        tbAnexos.Text += $";{openDialog.FileNames[i]}";
                    }
                }
            }
        }

        private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            DialogResult result = MessageBox.Show("Deseja realmente fechar?", Application.ProductName,
                MessageBoxButtons.YesNo, MessageBoxIcon.Question);

            if (result == DialogResult.No)
            {
                //cancela o evento de fechamento do formulário
                e.Cancel = true;
            }
        }

        private void BtnEnviar_Click(object sender, EventArgs e)
        {
            //verificando se todos os campos estão preenchidos
            if (string.IsNullOrEmpty(tbRemente.Text) |
                string.IsNullOrEmpty(tbSenhaRemetente.Text) |
                string.IsNullOrEmpty(tbDestinatario.Text) |
                string.IsNullOrEmpty(tbAssunto.Text) |
                string.IsNullOrEmpty(tbMsg.Text))
            {
                MessageBox.Show("Preencha todos os campos!", Application.ProductName,
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            //instânciando uma nova menssagem de email, passando o remetente e o destinatário
            using (var mailMsg = new MailMessage(new MailAddress(tbRemente.Text), new MailAddress(tbDestinatario.Text)))
            {
                mailMsg.Subject = tbAssunto.Text; //assunto do e-mail
                mailMsg.IsBodyHtml = false; //declarando que o corpo do e-mail não é em formato HTML
                mailMsg.BodyEncoding = Encoding.UTF8; //definindo que o e-mail seguirá o padrão de caracteres UTF-8
                mailMsg.Body = tbMsg.Text; //corpo do e-mail (mensagem a ser enviada)

                //verificando se o e-mail possui anexos
                if (cbAnexos.Checked)
                {
                    foreach (var item in tbAnexos.Text.Split(';'))
                    {
                        mailMsg.Attachments.Add(new Attachment(item));
                    }
                }

                try
                {
                    //instânciando um serviço SMTP para realizar o envio do e-mail
                    using (var smtp = new SmtpClient())
                    {
                        smtp.EnableSsl = true; //habilitando o certificado SSL
                        smtp.UseDefaultCredentials = false; //desativando as credenciais padrões

                        //passando as credenciais desejadas
                        smtp.Credentials = new NetworkCredential(tbRemente.Text, tbSenhaRemetente.Text);

                        smtp.Host = "smtp.office365.com"; //informando o host que será utilizado
                        smtp.Port = 587; //a porta que será utilizada
                        smtp.Send(mailMsg); //método que realizará o envio do e-mail
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Ocorreu um erro: {ex.Message}", Application.ProductName,
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }

            MessageBox.Show("E-mail enviado com sucesso!", Application.ProductName,
                MessageBoxButtons.OK);

            ClearFields();
        }

        private void ClearFields()
        {
            tbRemente.Clear();
            tbSenhaRemetente.Clear();
            tbDestinatario.Clear();
            tbAnexos.Clear();
            tbAssunto.Clear();
            tbMsg.Clear();
            cbAnexos.Checked = false;
        }
    }
}
