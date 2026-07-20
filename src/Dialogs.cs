using System.Windows.Forms;

namespace ExcelMerger
{
    /// <summary>
    /// Единый вход для диалогов приложения. Реальный вид — фирменный
    /// <see cref="MessageForm"/> (значок по типу, одна кнопка по центру, две — по краям).
    /// </summary>
    public static class Dialogs
    {
        /// <summary>Вопрос с кнопками Да/Нет и предупреждающим значком. true = Да.</summary>
        public static bool ConfirmWarning(IWin32Window owner, string title, string instruction, string content)
        {
            return MessageForm.ShowConfirm(owner, title, instruction, content);
        }

        public static void Info(IWin32Window owner, string title, string instruction, string content)
        {
            MessageForm.ShowMessage(owner, MessageForm.Kind.Info, title, instruction, content);
        }

        public static void Error(IWin32Window owner, string title, string instruction, string content)
        {
            MessageForm.ShowMessage(owner, MessageForm.Kind.Error, title, instruction, content);
        }

        /// <summary>Информационный диалог с кликабельной ссылкой под текстом.</summary>
        public static void InfoWithLink(IWin32Window owner, string title, string instruction, string content,
            string linkText, string url)
        {
            MessageForm.ShowMessage(owner, MessageForm.Kind.Info, title, instruction, content, linkText, url);
        }
    }
}
