using System;
using System.IO;
using System.Text;
using System.Windows.Forms;

namespace PluginSimulator.UI
{
    /// <summary>
    /// Redirects Console.Write/WriteLine output to a RichTextBox control.
    /// </summary>
    public class ConsoleRedirector : TextWriter
    {
        private readonly RichTextBox _output;
        private readonly StringBuilder _buffer = new StringBuilder();

        public ConsoleRedirector(RichTextBox output)
        {
            _output = output ?? throw new ArgumentNullException(nameof(output));
        }

        public override Encoding Encoding => Encoding.UTF8;

        public override void Write(char value)
        {
            _buffer.Append(value);
            if (value == '\n')
                Flush();
        }

        public override void Write(string value)
        {
            if (value == null) return;
            _buffer.Append(value);
            if (value.Contains("\n"))
                Flush();
        }

        public override void WriteLine(string value)
        {
            _buffer.AppendLine(value);
            Flush();
        }

        public override void Flush()
        {
            var text = _buffer.ToString();
            _buffer.Clear();

            if (string.IsNullOrEmpty(text)) return;

            if (_output.InvokeRequired)
            {
                _output.BeginInvoke(new Action(() => AppendText(text)));
            }
            else
            {
                AppendText(text);
            }
        }

        private void AppendText(string text)
        {
            _output.AppendText(text);
            _output.ScrollToCaret();
        }
    }
}
