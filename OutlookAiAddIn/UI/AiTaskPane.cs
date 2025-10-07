using System;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using OutlookAiAddIn.Services;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookAiAddIn.UI
{
    public partial class AiTaskPane : UserControl
    {
        private readonly OpenAIService _openAiService;
        private readonly OutlookContextService _contextService;
        private CancellationTokenSource _cts;
        private AiInteractionMode _mode = AiInteractionMode.SuggestedReply;

        public AiTaskPane(OpenAIService openAiService, OutlookContextService contextService)
        {
            InitializeComponent();

            _openAiService = openAiService;
            _contextService = contextService;

            cmbMode.Items.AddRange(new object[]
            {
                AiInteractionMode.SuggestedReply,
                AiInteractionMode.ImproveDraft,
                AiInteractionMode.Proofread
            });

            cmbMode.SelectedIndex = 0;
            UpdateLabelsForMode();
            UpdateUiState(false, "Pronto");
        }

        internal void LoadContext()
        {
            txtEmailBody.Text = _contextService.GetActiveMailBody();
        }

        internal void TriggerMode(AiInteractionMode mode)
        {
            if (!Equals(cmbMode.SelectedItem, mode))
            {
                cmbMode.SelectedItem = mode;
            }

            LoadContext();
            _ = GenerateForCurrentSettingsAsync();
        }

        internal async Task GenerateForCurrentSettingsAsync()
        {
            await GenerateAsync();
        }

        internal void HandleItemSend(object item, ref bool cancel)
        {
            CancelPendingRequest();
        }

        private async void btnGenerate_Click(object sender, EventArgs e)
        {
            await GenerateAsync();
        }

        private async Task GenerateAsync()
        {
            CancelPendingRequest();

            var context = txtEmailBody.Text;
            if (string.IsNullOrWhiteSpace(context))
            {
                LoadContext();
                context = txtEmailBody.Text;
            }

            if (string.IsNullOrWhiteSpace(context))
            {
                MessageBox.Show("Seleziona un'email o incolla del testo da analizzare.", "Assistente AI", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            _cts = new CancellationTokenSource();
            UpdateUiState(true, "Generazione in corso...");

            try
            {
                var suggestion = await _openAiService.GenerateAsync(_mode, context, txtNotes.Text, _cts.Token);
                txtOutput.Text = suggestion;
                UpdateUiState(false, "Suggerimento pronto.");
            }
            catch (OperationCanceledException)
            {
                UpdateUiState(false, "Richiesta annullata.");
            }
            catch (Exception ex)
            {
                UpdateUiState(false, $"Errore: {ex.Message}");
                MessageBox.Show(ex.Message, "Errore OpenAI", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                CancelPendingRequest(releaseOnly: true);
            }
        }

        private void btnLoadContext_Click(object sender, EventArgs e)
        {
            LoadContext();
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            CancelPendingRequest();
        }

        private void btnInsert_Click(object sender, EventArgs e)
        {
            var suggestion = txtOutput.Text;
            if (string.IsNullOrWhiteSpace(suggestion))
            {
                MessageBox.Show("Genera prima un suggerimento.", "Assistente AI", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            bool inserted = false;

            if (_mode == AiInteractionMode.SuggestedReply)
            {
                var mailItem = _contextService.EnsureReplyWindow();
                if (mailItem != null)
                {
                    if (mailItem.BodyFormat == Outlook.OlBodyFormat.olFormatHTML)
                    {
                        var html = ConvertPlainTextToHtmlParagraphs(suggestion);
                        inserted = _contextService.TryOverwriteDraft(html + mailItem.HTMLBody, html: true);
                    }
                    else
                    {
                        inserted = _contextService.TryOverwriteDraft(suggestion + Environment.NewLine + mailItem.Body, html: false);
                    }
                }
            }
            else
            {
                inserted = _contextService.TryInsertAtCursor(suggestion);
                if (!inserted)
                {
                    inserted = _contextService.TryOverwriteDraft(suggestion, html: false);
                }
            }

            if (inserted)
            {
                UpdateUiState(false, "Testo inserito nel messaggio.");
            }
            else
            {
                MessageBox.Show("Non è stato possibile inserire il testo automaticamente. Copialo e incollalo manualmente.", "Assistente AI", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void btnCopy_Click(object sender, EventArgs e)
        {
            var suggestion = txtOutput.Text;
            if (string.IsNullOrWhiteSpace(suggestion))
            {
                return;
            }

            try
            {
                Clipboard.SetText(suggestion);
                UpdateUiState(false, "Testo copiato negli appunti.");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Impossibile copiare il testo: {ex.Message}", "Assistente AI", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void cmbMode_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cmbMode.SelectedItem is AiInteractionMode newMode)
            {
                _mode = newMode;
                UpdateLabelsForMode();
            }
        }

        private void UpdateLabelsForMode()
        {
            switch (_mode)
            {
                case AiInteractionMode.SuggestedReply:
                    lblContext.Text = "Email originale";
                    lblNotes.Text = "Istruzioni aggiuntive per la risposta";
                    lblSuggestions.Text = "Risposta suggerita";
                    btnInsert.Text = "Crea risposta";
                    break;
                case AiInteractionMode.ImproveDraft:
                    lblContext.Text = "Bozza da migliorare";
                    lblNotes.Text = "Note sul tono desiderato";
                    lblSuggestions.Text = "Versione migliorata";
                    btnInsert.Text = "Sostituisci bozza";
                    break;
                case AiInteractionMode.Proofread:
                    lblContext.Text = "Testo da correggere";
                    lblNotes.Text = "Indicazioni (opzionale)";
                    lblSuggestions.Text = "Testo corretto";
                    btnInsert.Text = "Applica correzioni";
                    break;
            }
        }

        private void UpdateUiState(bool isBusy, string status)
        {
            btnGenerate.Enabled = !isBusy;
            btnCancel.Enabled = isBusy;
            btnLoadContext.Enabled = !isBusy;
            cmbMode.Enabled = !isBusy;
            txtNotes.ReadOnly = isBusy;
            txtEmailBody.ReadOnly = isBusy;
            btnInsert.Enabled = !isBusy && !string.IsNullOrWhiteSpace(txtOutput.Text);
            btnCopy.Enabled = !isBusy && !string.IsNullOrWhiteSpace(txtOutput.Text);
            lblStatus.Text = status;
        }

        private void txtOutput_TextChanged(object sender, EventArgs e)
        {
            if (_cts != null)
            {
                return;
            }

            btnInsert.Enabled = !string.IsNullOrWhiteSpace(txtOutput.Text);
            btnCopy.Enabled = !string.IsNullOrWhiteSpace(txtOutput.Text);
        }

        private void CancelPendingRequest(bool releaseOnly = false)
        {
            if (_cts == null)
            {
                return;
            }

            if (!releaseOnly)
            {
                _cts.Cancel();
            }

            _cts.Dispose();
            _cts = null;
        }

        private static string ConvertPlainTextToHtmlParagraphs(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
            {
                return string.Empty;
            }

            var encoded = System.Web.HttpUtility.HtmlEncode(text);
            encoded = encoded.Replace("\r\n", "<br/>").Replace("\n", "<br/>");
            return $"<div>{encoded}</div>";
        }
    }
}
