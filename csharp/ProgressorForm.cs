/*
 * Создано в SharpDevelop.
 * Пользователь: egor
 * Дата: 01.09.2017
 * Время: 14:19
 * 
 * Для изменения этого шаблона используйте меню "Инструменты | Параметры | Кодирование | Стандартные заголовки".
 */
using System;
using System.Drawing;
using System.Windows.Forms;

namespace badge
{
	/// <summary>
	/// Description of ProgressorForm.
	/// </summary>
	public partial class ProgressorForm : Form
	{
		private int allCount;
		public ProgressorForm(int all)
		{
			//
			// The InitializeComponent() call is required for Windows Forms designer support.
			//
			InitializeComponent();
			
			this.allCount = all;
			this.label1.Text = "0%";
			this.progressBar1.Visible = true;
			this.progressBar1.Maximum = all;
			this.progressBar1.Step = 1;
			this.progressBar1.Minimum = 0;
			this.progressBar1.Value = 0;
		}
		
		public void DoStep() {
			this.progressBar1.PerformStep();
			this.label1.Text = ((int)(this.progressBar1.Value * 100 / this.progressBar1.Maximum)).ToString() + "%";
		}
		
		public void Quit() {
			this.Dispose();
		}
	}
}
