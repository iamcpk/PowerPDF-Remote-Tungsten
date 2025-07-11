/* Created by Aleah Hassabo
Target Framework used .Net 9.0
C# version 13.0 */
using System;
using System.Diagnostics;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;
using WindowsInput;

namespace PowerPDF_Remote_Tungsten
{
    public partial class Form1 : Form
    {
        [DllImport("user32.dll")]
        private static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        private Button button1, button2, button3, button4, button5, button6, button7, button8, button9;
        private PictureBox logoPictureBox, pdfLogoPictureBox; // Add a PictureBox for the logo
        private ToolTip ToolTip1, ToolTip2, ToolTip3, ToolTip4, ToolTip5, ToolTip6, ToolTip7, ToolTip8, ToolTip9;
        public Form1()
        {
            InitializeComponent();
            this.BackColor = ColorTranslator.FromHtml("#002854");
            this.Size = new Size(700, 700);

            CreateLogo();
            this.Resize += new EventHandler(Form1_Resize);
            CreateButtons();
            CenterButtons();
            this.Resize += Form1_Resize;
        }

        private void CreateLogo()
        {
            logoPictureBox = new PictureBox();
            logoPictureBox.Size = new Size(140, 60); // Set the size of the logo
            logoPictureBox.Image = System.Drawing.Image.FromFile(@".\images\tungstenwhitelogo.png"); // Load your logo image from images
            logoPictureBox.SizeMode = PictureBoxSizeMode.Zoom; // Adjust the size mode
            logoPictureBox.Location = new Point(this.ClientSize.Width - logoPictureBox.Width - 10, 10);
            this.Controls.Add(logoPictureBox);

            pdfLogoPictureBox = new PictureBox();
            pdfLogoPictureBox.Size = new Size(160, 60); // Set the size of the logo
            pdfLogoPictureBox.Image = System.Drawing.Image.FromFile(@".\images\powerpdfadvanvcedwhite.png"); // Load your logo image
            pdfLogoPictureBox.SizeMode = PictureBoxSizeMode.Zoom; // Adjust the size mode
            pdfLogoPictureBox.Location = new Point(15, 10);
            this.Controls.Add(pdfLogoPictureBox);
        }

        private void CreateButtons()
        {


            button1 = new Button();
            button1.BackColor = Color.White;
            button1.Size = new System.Drawing.Size(100, 50);
            ToolTip1 = new ToolTip();
            ToolTip1.SetToolTip(button1, "Text comparator for 2 PDF's");
            // button1.Text = "Text Comparator"; // Set button text
            button1.Image = System.Drawing.Image.FromFile(@".\images\B1.png");
            button1.ImageAlign = ContentAlignment.MiddleCenter;
            button1.Click += (sender, e) => Button_Click(1); // Assign click event
            this.Controls.Add(button1);

            button2 = new Button();
            button2.BackColor = Color.White;
            button2.Size = new System.Drawing.Size(100, 50);
            ToolTip2 = new ToolTip();
            ToolTip2.SetToolTip(button2, "Combines your choice of PDF's");
            // button2.Text = "Combine All Open PDF's"; // Set button text
            button2.Image = System.Drawing.Image.FromFile(@".\images\B2.png");
            button2.ImageAlign = ContentAlignment.MiddleCenter;
            button2.Click += (sender, e) => Button_Click(2); // Assign click event
            this.Controls.Add(button2);

            button3 = new Button();
            button3.BackColor = Color.White;
            button3.Size = new System.Drawing.Size(100, 50);
            ToolTip3 = new ToolTip();
            ToolTip3.SetToolTip(button3, "Rotates a page and or the entire PDF");
            //button3.Text = "Button 3"; // Set button text
            button3.Image = System.Drawing.Image.FromFile(@".\images\B3.png");
            button3.ImageAlign = ContentAlignment.MiddleCenter;
            button3.Click += (sender, e) => Button_Click(3); // Assign click event
            this.Controls.Add(button3);

            button4 = new Button();
            button4.BackColor = Color.White;
            button4.Size = new System.Drawing.Size(100, 50);
            ToolTip4 = new ToolTip();
            ToolTip4.SetToolTip(button4, "Convert PDF to an Excel Sheet");
            //button4.Text = "Button 4"; // Set button text
            button4.Image = System.Drawing.Image.FromFile(@".\images\B4.png");
            button4.ImageAlign = ContentAlignment.MiddleCenter;
            button4.Click += (sender, e) => Button_Click(4); // Assign click event
            this.Controls.Add(button4);

            button5 = new Button();
            button5.BackColor = Color.White;
            button5.Size = new System.Drawing.Size(100, 50);
            ToolTip5 = new ToolTip();
            ToolTip5.SetToolTip(button5, "Share PDF as an e-mail");
            //button5.Text = "Button 5"; // Set button text
            button5.Image = System.Drawing.Image.FromFile(@".\images\B5.png");
            button5.ImageAlign = ContentAlignment.MiddleCenter;
            button5.Click += (sender, e) => Button_Click(5); // Assign click event
            this.Controls.Add(button5);

            button6 = new Button();
            button6.BackColor = Color.White;
            button6.Size = new System.Drawing.Size(100, 50);
            ToolTip6 = new ToolTip();
            ToolTip6.SetToolTip(button6, "Create a new portfolio");
            //button6.Text = "Button 6"; // Set button text
            button6.Image = System.Drawing.Image.FromFile(@".\images\B6.png");
            button6.ImageAlign = ContentAlignment.MiddleCenter;
            button6.Click += (sender, e) => Button_Click(6); // Assign click event
            this.Controls.Add(button6);

            button7 = new Button();
            button7.BackColor = Color.White;
            button7.Size = new System.Drawing.Size(100, 50);
            ToolTip7 = new ToolTip();
            ToolTip7.SetToolTip(button7, "Save file as");
            // button7.Text = "Button 7"; // Set button text
            button7.Image = System.Drawing.Image.FromFile(@".\images\B7.png");
            button7.ImageAlign = ContentAlignment.MiddleCenter;
            button7.Click += (sender, e) => Button_Click(7); // Assign click event
            this.Controls.Add(button7);

            button8 = new Button();
            button8.BackColor = Color.White;
            button8.Size = new System.Drawing.Size(100, 50);
            ToolTip8 = new ToolTip();
            ToolTip8.SetToolTip(button8, "Convert PDF to a PowerPoint");
            // button8.Text = "Button 8"; // Set button text
            button8.Image = System.Drawing.Image.FromFile(@".\images\B8.png");
            button8.ImageAlign = ContentAlignment.MiddleCenter;
            button8.Click += (sender, e) => Button_Click(8); // Assign click event
            this.Controls.Add(button8);

            button9 = new Button();
            button9.BackColor = Color.White;
            button9.Size = new System.Drawing.Size(100, 50);
            ToolTip9 = new ToolTip();
            ToolTip9.SetToolTip(button9, "Flatten Power PDF");
            // button9.Text = "Button 9"; // Set button text
            button9.Image = System.Drawing.Image.FromFile(@".\images\B9.png");
            button9.ImageAlign = ContentAlignment.MiddleCenter;
            button9.Click += (sender, e) => Button_Click(9); // Assign click event
            this.Controls.Add(button9);

        }

        private void CenterButtons()
        {
            int cols = 3;
            int rows = 3;
            int bottomPadding = 20; // Set the desired padding at the bottom

            // Calculate button size based on the form's client size, accounting for the logo height and bottom padding
            int buttonWidth = (this.ClientSize.Width - (cols + 1) * 10) / cols;
            int buttonHeight = (this.ClientSize.Height - (rows + 1) * 10 - logoPictureBox.Height - bottomPadding) / rows; // Subtract logo height and bottom padding

            Button[,] buttons = {
        { button1, button2, button3 },
        { button4, button5, button6 },
        { button7, button8, button9 }
    };

            for (int row = 0; row < rows; row++)
            {
                for (int col = 0; col < cols; col++)
                {
                    var button = buttons[row, col];
                    button.Size = new Size(buttonWidth, buttonHeight);
                    button.Location = new Point(10 + col * (buttonWidth + 10), logoPictureBox.Bottom + 10 + row * (buttonHeight + 10)); // Position below the logo
                }
            }
        }

        private void Button_Click(int buttonNumber)
        {
            var myPowerPDF = new PDFPlus.App();
            var myDoc = new PDFPlus.DVDoc();
            var simulator = new InputSimulator();

            [DllImport("user32.dll")]
            static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

            [DllImport("user32.dll")]
            static extern bool SetForegroundWindow(IntPtr hWnd);

            IntPtr zero = IntPtr.Zero;

            switch (buttonNumber)
            {
                case 1: //autocompare pdfs
                    new Thread(() =>
                    {
                        Thread.CurrentThread.IsBackground = true;
                        myPowerPDF.MenuItemExecute("AutoCompare");
                    }).Start();

                    for (int i = 0; (i < 50) && (zero == IntPtr.Zero); i++)
                    {
                        zero = FindWindow(null, "Combine Files");
                        if (zero != IntPtr.Zero) break; // Exit loop early if window is found
                    }
                    if (zero != IntPtr.Zero)
                    {
                        SetForegroundWindow(zero);
                        SendKeys.SendWait("{ENTER}");
                        SendKeys.Flush();
                    }
                    break;

                case 2: //combining files
                    new Thread(() =>
                    {
                        Thread.CurrentThread.IsBackground = true;
                        myPowerPDF.MenuItemExecute("CombinAllOpenFile");
                    }).Start();

                    for (int i = 0; (i < 50) && (zero == IntPtr.Zero); i++)
                    {
                        zero = FindWindow(null, "Combine Files");
                        if (zero != IntPtr.Zero) break; // Exit loop early if window is found
                    }
                    if (zero != IntPtr.Zero)
                    {
                        SetForegroundWindow(zero);
                        SendKeys.SendWait("{ENTER}");
                        SendKeys.Flush();
                    }
                    break;

                case 3: //rotate entire pdf
                    new Thread(() =>
                    {
                        Thread.CurrentThread.IsBackground = true;
                        myPowerPDF.MenuItemExecute("RotatePage");
                    }).Start();

                    for (int i = 0; (i < 50) && (zero == IntPtr.Zero); i++)
                    {
                        zero = FindWindow(null, "Rotate Pages");
                        if (zero != IntPtr.Zero) break; // Exit loop early if window is found
                    }

                    if (zero != IntPtr.Zero)
                    {
                        SetForegroundWindow(zero);
                        SendKeys.SendWait("{ENTER}");
                        SendKeys.Flush();
                    }
                    break;

                case 4: //save file as an excel sheet
                    new Thread(() =>
                    {
                        Thread.CurrentThread.IsBackground = true;
                        myPowerPDF.MenuItemExecute("SaveAsExcel");
                    }).Start();

                    for (int i = 0; (i < 50) && (zero == IntPtr.Zero); i++)
                    {
                        zero = FindWindow(null, "Convert Pages");
                        if (zero != IntPtr.Zero) break; // Exit loop early if window is found
                    }

                    if (zero != IntPtr.Zero)
                    {
                        SetForegroundWindow(zero);
                        SendKeys.SendWait("{ENTER}");
                        SendKeys.Flush();
                    }
                    break;

                case 5: //share as email
                    new Thread(() =>
                    {
                        Thread.CurrentThread.IsBackground = true;
                        myPowerPDF.MenuItemExecute("Email");
                    }).Start();

                    for (int i = 0; (i < 50) && (zero == IntPtr.Zero); i++)
                    {
                        Thread.Sleep(100);
                        zero = FindWindow(null, "E-mail");
                        if (zero != IntPtr.Zero) break; // Exit loop early if window is found
                    }

                    if (zero != IntPtr.Zero)
                    {
                        SetForegroundWindow(zero);
                        SendKeys.SendWait("{ENTER}");
                        SendKeys.Flush();
                    }
                    break;

                case 6: //create a new portfolio
                    new Thread(() =>
                    {
                        Thread.CurrentThread.IsBackground = true;
                        myPowerPDF.MenuItemExecute("CreatePortfolio");
                    }).Start();

                    //Thread.Sleep(100);

                    //for (int i = 0; (i < 50) && (zero == IntPtr.Zero); i++)
                    //{
                    //    Thread.Sleep(100);
                    //    zero = FindWindow(null, "");
                    //    if (zero != IntPtr.Zero) break; // Exit loop early if window is found
                    //}

                    //if (zero != IntPtr.Zero)
                    //{
                    //    SetForegroundWindow(zero);
                    //    SendKeys.SendWait("{ENTER}");
                    //    SendKeys.Flush();
                    //}
                    break;

                case 7: //save file as
                    new Thread(() =>
                    {
                        Thread.CurrentThread.IsBackground = true;
                        myPowerPDF.MenuItemExecute("SaveAs");
                    }).Start();

                    for (int i = 0; (i < 50) && (zero == IntPtr.Zero); i++)
                    {
                        zero = FindWindow(null, "Save As");
                        if (zero != IntPtr.Zero) break; // Exit loop early if window is found
                    }

                    if (zero != IntPtr.Zero)
                    {
                        SetForegroundWindow(zero);
                        SendKeys.SendWait("{ENTER}");
                        SendKeys.Flush();
                    }
                    break;

                case 8: //save file as a powerpoint
                    new Thread(() =>
                    {
                        Thread.CurrentThread.IsBackground = true;
                        myPowerPDF.MenuItemExecute("SaveAsPPT");
                    }).Start();

                    for (int i = 0; (i < 50) && (zero == IntPtr.Zero); i++)
                    {
                        zero = FindWindow(null, "Convert Pages");
                        if (zero != IntPtr.Zero) break; // Exit loop early if window is found
                    }
                    if (zero != IntPtr.Zero)
                    {
                        SetForegroundWindow(zero);
                        SendKeys.SendWait("{ENTER}");
                        SendKeys.Flush();
                    }
                    break;

                case 9: //Flatten File
                    new Thread(() =>
                    {
                        Thread.CurrentThread.IsBackground = true;
                        myPowerPDF.MenuItemExecute("Flatten");
                    }).Start();

                    for (int i = 0; (i < 50) && (zero == IntPtr.Zero); i++)
                    {
                        zero = FindWindow(null, "Flatten");
                        if (zero != IntPtr.Zero) break; // Exit loop early if window is found
                    }

                    if (zero != IntPtr.Zero)
                    {
                        SetForegroundWindow(zero);
                        SendKeys.SendWait("{ENTER}");
                        SendKeys.Flush();
                    }
                    break;



                default:
                    // Default behavior for other buttons
                    MessageBox.Show($"Button {buttonNumber} clicked!");
                    break;
            }
        }

        private void Form1_Resize(object sender, EventArgs e)
        {
            CenterButtons(); // Recalculate positions on resize

            logoPictureBox.Location = new Point(this.ClientSize.Width - logoPictureBox.Width - 10, 10);
            logoPictureBox.Size = new Size(this.ClientSize.Width / 5, this.ClientSize.Height / 10);

            pdfLogoPictureBox.Location = new Point(10, 10); // Adjust as needed
            pdfLogoPictureBox.Size = new Size(this.ClientSize.Width / 5, this.ClientSize.Height / 10);
        }
        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void Form1_Load_1(object sender, EventArgs e)
        {

        }
    }
}
