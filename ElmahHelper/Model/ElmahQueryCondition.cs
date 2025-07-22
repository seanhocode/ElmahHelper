using System;
using System.Windows.Forms;

namespace ElmahHelper.Model
{
    public class ElmahQueryCondition
    {
        public DateTimePicker StartTimePicker { get; set; }
        public DateTimePicker EndTimePicker { get; set; }
        public TextBox FileNameTextBox { get; set; }
        public TextBox MessageTextBox { get; set; }
        public TextBox DetailTextBox { get; set; }
        public DateTime StartTime
        {
            get => StartTimePicker.Value;
            set => StartTimePicker.Value = value;
        }
        public DateTime EndTime
        {
            get => EndTimePicker.Value;
            set => EndTimePicker.Value = value;
        }
        public string FileName
        {
            get => FileNameTextBox.Text;
            set => FileNameTextBox.Text = value;
        }
        public string Message
        {
            get => MessageTextBox.Text;
            set => MessageTextBox.Text = value;
        }
        public string Detail
        {
            get => DetailTextBox.Text;
            set => DetailTextBox.Text = value;
        }
    }
}
