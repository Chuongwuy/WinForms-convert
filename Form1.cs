using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection.Emit;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Button;
using OfficeOpenXml;

namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {
        private string selectedFilePath = "";
<<<<<<< HEAD
        //string outputMixer = @"C:\Users\HP\OneDrive\Documents\Academic Documents\MACOM\output_Mixer";
        //string outputDoubler = @"C:\Users\HP\OneDrive\Documents\Academic Documents\MACOM\output_Doubler";
        string outputMixer;
        string outputDoubler;
=======
        string outputMixer = @"C:\Users\HP\OneDrive\Documents\Academic Documents\MACOM\output_Mixer";
        string outputDoubler = @"C:\Users\HP\OneDrive\Documents\Academic Documents\MACOM\output_Doubler";
>>>>>>> 2742ae2f9991af8eb6b861e2a4bb4cf034f1f554

        void UpdateStatus(string message)
        {
            // Ví dụ cập nhật một Label để hiển thị trạng thái
            resultlabel.Text = message;
        }

        public Form1()
        {
            InitializeComponent();
            ExcelPackage.License.SetNonCommercialPersonal("HUY");    //License NET 4.8


            // Gán sự kiện cho  các nút nhấn 
            browsebutton.Click += (s, e) => BrowseFile();
<<<<<<< HEAD

            outputbutton.Click += (s, e) => Browseoutput();


=======
>>>>>>> 2742ae2f9991af8eb6b861e2a4bb4cf034f1f554
            cancelbutton.Click += (s, e) => Application.Exit();
            generatebutton.Click += (s, e) => GenerateFiles();

        }

        // KHÔNG ĐỤNG VÀO PHẦN DƯỚI NÀY 
        //--------------------------------------------------------------
        private void Form1_Load(object sender, EventArgs e)
        {

        }
<<<<<<< HEAD
        private void label1_Click_1(object sender, EventArgs e)
        {

        }
=======

>>>>>>> 2742ae2f9991af8eb6b861e2a4bb4cf034f1f554
        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {

        }
        //--------------------------------------------------------------


        private void BrowseFile()
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm|All Files|*.*";
                openFileDialog.Title = "Select Excel File";

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    selectedFilePath = openFileDialog.FileName;
                    textbox.Text = selectedFilePath;             // write the filepath into textbox 
                    UpdateStatus("File selected: " + Path.GetFileName(selectedFilePath));
                }
            }
        }



<<<<<<< HEAD
        private void Browseoutput()
        {
            using (FolderBrowserDialog folderBrowserDialog = new FolderBrowserDialog())
            {
                folderBrowserDialog.Description = "Select Output Folder";
                if (folderBrowserDialog.ShowDialog() == DialogResult.OK)
                {
                    // Lưu đường dẫn thư mục đã chọn
                    outputMixer = folderBrowserDialog.SelectedPath;
                    textbox2.Text = outputMixer;  
                    outputDoubler = folderBrowserDialog.SelectedPath;
                    UpdateStatus("Folder selected: " + outputMixer);
                }
            }
        }





=======
>>>>>>> 2742ae2f9991af8eb6b861e2a4bb4cf034f1f554
        private void DoublerCheckbox(object sender, EventArgs e)
        {
            if (mixercheckbox.Checked)
            {
                // ✅ CheckBox được tick
                // TODO: Xử lý theo hướng bạn muốn
            }
            else
            {
                // ❌ CheckBox chưa được tick
                // TODO: Xử lý theo hướng khác
            }
        }



        private void GenerateFiles()
        {
            bool doublerSelected = doublercheckbox.Checked;
            bool mixerSelected = mixercheckbox.Checked;

            if (string.IsNullOrEmpty(selectedFilePath))
            {
<<<<<<< HEAD
                MessageBox.Show($"Please select a input Excel file first", "Reminder", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            if (string.IsNullOrEmpty(outputMixer))
            {
                MessageBox.Show($"Please select an output folder", "Reminder", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }


=======
                MessageBox.Show($"Please select a file first", "Reminder", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

>>>>>>> 2742ae2f9991af8eb6b861e2a4bb4cf034f1f554
            if (!doublerSelected && !mixerSelected)
            {
                MessageBox.Show($"Please choose at least one model in checkbox", "Reminder", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            try
            {
                // Kiểm tra file tồn tại
                if (!File.Exists(selectedFilePath))
                {
                    MessageBox.Show($"File not found at {selectedFilePath} ", "Reminder", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                // Tạo thư mục output nếu chưa tồn tại
                //string outputDir = Path.Combine(Path.GetDirectoryName(selectedFilePath), "Output");
                //Directory.CreateDirectory(outputDir);

                // Xử lý cho Doubler
                if (doublerSelected)
                {

                    ProcessModelDoubler();   //xử lý Doubler và xuất output vào thư mục outputDoubler
                }

                // Xử lý cho Mixer
                if (mixerSelected)
                {
                    ProcessModelMixer();
                }

                //UpdateStatus("Processing completed successfully!");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }


        // -------------------------------------------------------------XỬ LÝ MIXER ---------------------------------------------------------------
        private void ProcessModelMixer()
        {
            UpdateStatus($"Processing Mixer model file...");
            System.Threading.Thread.Sleep(1500); // Giả lập thời gian xử lý
            try
            {
                // create file ini
                string timestamp = DateTime.Now.ToString("HH-mm-ss__dd-MM-yyyy");
<<<<<<< HEAD
                string outputFileName = $"config_Mixer_{timestamp}.ini";
=======
                string outputFileName = $"config_{timestamp}.ini";
>>>>>>> 2742ae2f9991af8eb6b861e2a4bb4cf034f1f554
                string output = Path.Combine(outputMixer, outputFileName);
                // StringBuilder for ini results
                var outputBuilder = new StringBuilder();
                var fileInfo = new FileInfo(selectedFilePath);
                using (var package = new ExcelPackage(fileInfo))
                {
                    // check worksheet 
                    var worksheet = package.Workbook.Worksheets[0];
                    if (worksheet == null)
                    {
                        Console.WriteLine("Can not find any worksheet in file Excel.");
<<<<<<< HEAD
                        MessageBox.Show("Can not find any worksheet in file Excel.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
=======
>>>>>>> 2742ae2f9991af8eb6b861e2a4bb4cf034f1f554
                        return;
                    }

                    {
                        outputBuilder.AppendLine($"[INFO]\n");
                        //--------------------------------------NumCLStep= -------------------------------------------
                        int numCols = 0;
                        if (int.TryParse(worksheet.Cells["B1"].Value?.ToString(), out int n_value))
                        {
                            numCols = n_value;
                        }

                        int numRows = 0;
                        if (int.TryParse(worksheet.Cells["A2"].Value?.ToString(), out int m_value))
                        {
                            numRows = m_value;
                        }

                        if (numCols <= 0 || numRows <= 0)
                        {
                            Console.WriteLine("Value unvalid.");
                            outputBuilder.AppendLine($"Value unvalid.");
                            File.WriteAllText(outputMixer, outputBuilder.ToString());
<<<<<<< HEAD
                            MessageBox.Show("Values at B1 and A2 are unvalid or null. Model Mixer can not be converted", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
=======
>>>>>>> 2742ae2f9991af8eb6b861e2a4bb4cf034f1f554
                            return;
                        }


                        // Check data 
                        // Hàng bắt đầu từ 4, cột bắt đầu từ 2 (cột B)
                        int startRow = 4;
                        int startCol = 2;
                        int step = 0;
                        // Check matrix size 
                        for (int row = startRow; row < startRow + numRows; row++)
                        {
                            for (int col = startCol; col < startCol + numCols; col++)
                            {
                                // Lấy giá trị của ô hiện tại
                                var cellValue = worksheet.Cells[row, col].Value;
                                // 4. Kiểm tra xem ô có dữ liệu hay không (không rỗng, không null)
                                if (cellValue != null && !string.IsNullOrWhiteSpace(cellValue.ToString()))
                                {
                                    step++;
                                }
                            }
                        }

                        outputBuilder.AppendLine($"NumCLStep={step}");
                        //------------------------------------NumToiStep ------------------------------------------------
                        int numRows1 = 0;
                        if (int.TryParse(worksheet.Cells["A21"].Value?.ToString(), out int h_value))
                        {
                            numRows1 = h_value;
                        }
                        if (numRows1 <= 0)
                        {
                            Console.WriteLine("Value unvalid.");
<<<<<<< HEAD
                            outputBuilder.AppendLine($"Values at B1 and A2 are unvalid or null.");
                            File.WriteAllText(outputMixer, outputBuilder.ToString());
                            MessageBox.Show("Values at B1 and A2 are unvalid or null. Model Mixer can not be converted", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
=======
                            outputBuilder.AppendLine($"Value unvalid.");
                            File.WriteAllText(outputMixer, outputBuilder.ToString());
>>>>>>> 2742ae2f9991af8eb6b861e2a4bb4cf034f1f554
                            return;
                        }
                        int startRow1 = 22;
                        int step1 = 0;
                        // 3. Lặp qua từng ô trong ma trận đã xác định
                        for (int row1 = startRow1; row1 < startRow1 + numRows1; row1++)
                        {
                            // Lấy giá trị của ô hiện tại
                            var cellValue1 = worksheet.Cells[row1, 1].Value;
                            // 4. Kiểm tra xem ô có dữ liệu hay không (không rỗng, không null)
                            if (cellValue1 != null && !string.IsNullOrWhiteSpace(cellValue1.ToString()))
                            {
                                step1++;
                            }
                        }

                        outputBuilder.AppendLine($"NumToiStep={step1}");
                        // -----------------------------------------NumIsoStep -----------------------------------------------
                        int numRows2 = 0;
                        if (int.TryParse(worksheet.Cells["A41"].Value?.ToString(), out int f_value))
                        {
                            numRows2 = f_value;
                        }
                        if (numRows2 <= 0)
                        {
                            Console.WriteLine("Value unvalid.");
                            outputBuilder.AppendLine($"Value unvalid.");
                            File.WriteAllText(outputMixer , outputBuilder.ToString());
<<<<<<< HEAD
                            MessageBox.Show("Values at B1 and A2 are unvalid or null. Model Mixer can not be converted", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
=======
>>>>>>> 2742ae2f9991af8eb6b861e2a4bb4cf034f1f554
                            return;
                        }
                        int startRow2 = 42;
                        int step2 = 0;
                        // 3. Lặp qua từng ô trong ma trận đã xác định
                        for (int row2 = startRow2; row2 < startRow2 + numRows2; row2++)
                        {
                            // Lấy giá trị của ô hiện tại
                            var cellValue2 = worksheet.Cells[row2, 1].Value;
                            // 4. Kiểm tra xem ô có dữ liệu hay không (không rỗng, không null)
                            if (cellValue2 != null && !string.IsNullOrWhiteSpace(cellValue2.ToString()))
                            {
                                step2++;
                            }
                        }
                
                        outputBuilder.AppendLine($"NumIsoStep={step2}");

                        // ---------------------------------------NumDblStep----------------------------------------
             
                        outputBuilder.AppendLine($"NumDblStep=0 ");
                        // ------------------------------------------Comment ---------------------------------------
                        var cellValue4 = worksheet.Cells["F2"].Value;
                        if (cellValue4 != null && !string.IsNullOrWhiteSpace(cellValue4.ToString()))
                        {
         
                            outputBuilder.AppendLine($"Comment= {cellValue4}");
                        }
                        else
                        {
 
                            outputBuilder.AppendLine($"Comment= ");
                        }
                        // ------------------------------------------CAl_date---------------------------------------

                        string date = DateTime.Now.ToString("dd-MM-yyyy");

        
                        outputBuilder.AppendLine($"CAL_DATE= {date}");

                    }    // INFO 
                    {
                        Console.WriteLine($"\n");
                        Console.WriteLine($"[CL_RF_SRC1]\n");
                        outputBuilder.AppendLine($"\n");
                        outputBuilder.AppendLine($"[CL_RF_SCR1]\n");
                        // ---------------------- FREQUENCY ------------------------------------

                        // 1. Đọc số lượng cột (n) từ ô B1 và số lượng hàng (m) từ ô A2
                        // Chuyển đổi giá trị ô sang số nguyên, nếu không thành công thì dùng giá trị mặc định là 0.
                        int numCols = 0;
                        if (int.TryParse(worksheet.Cells["B1"].Value?.ToString(), out int n_value))
                        {
                            numCols = n_value;
                        }

                        int numRows = 0;
                        if (int.TryParse(worksheet.Cells["A2"].Value?.ToString(), out int m_value))
                        {
                            numRows = m_value;
                        }

                        if (numCols <= 0 || numRows <= 0)
                        {
                            Console.WriteLine("Value unvalid.");
                            outputBuilder.AppendLine($"Value unvalid.");
                            File.WriteAllText(outputMixer, outputBuilder.ToString());
<<<<<<< HEAD
                            MessageBox.Show("Values at B1 and A2 are unvalid or null. Model Mixer can not be converted", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
=======
>>>>>>> 2742ae2f9991af8eb6b861e2a4bb4cf034f1f554
                            return;
                        }


                        // 2. Xác định vùng dữ liệu (ma trận màu vàng) để quét
                        // Hàng bắt đầu từ 4, cột bắt đầu từ 2 (cột B)
                        int startRow = 4;
                        int startCol = 2;
                        int counter = 1;


                        // Hàm xử lý định dạng số (thay thế 6 số 0 cuối bằng 'e6')
                        string FormatNumber(string number)
                        {
                            if (string.IsNullOrEmpty(number))
                                return "N/A";

                            // Kiểm tra nếu chuỗi kết thúc bằng 6 số 0
                            if (number.EndsWith("000000"))
                            {
                                return number.Substring(0, number.Length - 6) + "e6";
                            }
                            return number;
                        }


                        // 3. Lặp qua từng ô trong ma trận đã xác định
                        for (int row = startRow; row < startRow + numRows; row++)
                        {
                            for (int col = startCol; col < startCol + numCols; col++)
                            {
                                // Lấy giá trị của ô hiện tại
                                var cellValue = worksheet.Cells[row, col].Value;
                                // 4. Kiểm tra xem ô có dữ liệu hay không (không rỗng, không null)
                                if (cellValue != null && !string.IsNullOrWhiteSpace(cellValue.ToString()))
                                {

                                    // Lấy giá trị input từ cột A cùng hàng
                                    var inputValue = worksheet.Cells[row, 1].Value?.ToString();
                                    // Xử lý định dạng số (thay thế 6 số 0 bằng 'e6')
                                    string formattedInput = FormatNumber(inputValue);
                                    //kết quả terminal 
                                    Console.WriteLine($"Freq{counter}={formattedInput}");
                                    outputBuilder.AppendLine($"Freq{counter}={formattedInput}");
                                    counter++;
                                }
                            }
                        }
                    }    // Input Frequency [CL_RF_SRC1]    Conversion Loss
                    {
                        Console.WriteLine($"\n");
                        outputBuilder.AppendLine($"\n");

                        //------------------------------------- POWER ------------------------------------------ 
                        int numPowerInput = 0;
                        if (int.TryParse(worksheet.Cells["D2"].Value?.ToString(), out int p_value))
                        {
                            numPowerInput = p_value;
                        }

                        int numCols = 0;
                        if (int.TryParse(worksheet.Cells["B1"].Value?.ToString(), out int n_value))
                        {
                            numCols = n_value;
                        }

                        int numRows = 0;
                        if (int.TryParse(worksheet.Cells["A2"].Value?.ToString(), out int m_value))
                        {
                            numRows = m_value;
                        }

                        if (numCols <= 0 || numRows <= 0)
                        {
                            Console.WriteLine("Value unvalid.");
                            outputBuilder.AppendLine($"Value unvalid");
                            File.WriteAllText(outputMixer, outputBuilder.ToString());
<<<<<<< HEAD
                            MessageBox.Show("Values at B1 and A2 are unvalid or null. Model Mixer can not be converted", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
=======
>>>>>>> 2742ae2f9991af8eb6b861e2a4bb4cf034f1f554
                            return;
                        }

                        int startRow = 4;
                        int startCol = 2;
                        int counter = 1;
                        for (int row = startRow; row < startRow + numRows; row++)
                        {
                            for (int col = startCol; col < startCol + numCols; col++)
                            {
                                // Lấy giá trị của ô hiện tại
                                var cellValue = worksheet.Cells[row, col].Value;
                                // 4. Kiểm tra xem ô có dữ liệu hay không (không rỗng, không null)
                                if (cellValue != null && !string.IsNullOrWhiteSpace(cellValue.ToString()))
                                {

                                    // Lấy giá trị input từ cột A cùng hàng
                                    //var inputValue = worksheet.Cells[row, 1].Value?.ToString();
                                    // Xử lý định dạng số (thay thế 6 số 0 bằng 'e6')
                                    //string formattedInput = FormatNumber(inputValue);
                                    //kết quả terminal 
                                    Console.WriteLine($"Power{counter}={numPowerInput}");
                                    outputBuilder.AppendLine($"Power{counter}={numPowerInput}");
                                    counter++;
                                }
                            }
                        }
                    }    // POWER [CL_RF_SRC1]
                    {
                        Console.WriteLine($"\n");
                        outputBuilder.AppendLine($"\n");
                        // --------------------------SETPOWER ------------------------ 
                        int numCols = 0;
                        if (int.TryParse(worksheet.Cells["B1"].Value?.ToString(), out int n_value))
                        {
                            numCols = n_value;
                        }

                        int numRows = 0;
                        if (int.TryParse(worksheet.Cells["A2"].Value?.ToString(), out int m_value))
                        {
                            numRows = m_value;
                        }

                        if (numCols <= 0 || numRows <= 0)
                        {
                            Console.WriteLine("Value unvalid.");
                            outputBuilder.AppendLine($"Value unvalid");
                            File.WriteAllText(outputMixer, outputBuilder.ToString());
<<<<<<< HEAD
                            MessageBox.Show("Values at B1 and A2 are unvalid or null. Model Mixer can not be converted", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
=======
>>>>>>> 2742ae2f9991af8eb6b861e2a4bb4cf034f1f554
                            return;
                        }

                        int startRow = 4;
                        int startCol = 2;
                        int counter = 1;
                        for (int row = startRow; row < startRow + numRows; row++)
                        {
                            for (int col = startCol; col < startCol + numCols; col++)
                            {
                                // Lấy giá trị của ô hiện tại
                                var cellValue = worksheet.Cells[row, col].Value;
                                // 4. Kiểm tra xem ô có dữ liệu hay không (không rỗng, không null)
                                if (cellValue != null && !string.IsNullOrWhiteSpace(cellValue.ToString()))
                                {

                                    // Lấy giá trị input từ cột A cùng hàng
                                    //var inputValue = worksheet.Cells[row, 1].Value?.ToString();
                                    // Xử lý định dạng số (thay thế 6 số 0 bằng 'e6')
                                    //string formattedInput = FormatNumber(inputValue);
                                    //kết quả terminal 
                                    Console.WriteLine($"SetPower{counter}=0");
                                    outputBuilder.AppendLine($"SetPower{counter}=0");
                                    counter++;
                                }
                            }
                        }
                    }    // SETPOWER [CL_RF_SRC1]
                    {
                        Console.WriteLine($"\n");
                        Console.WriteLine($"[CL_LO_SRC]\n");
                        outputBuilder.AppendLine($"\n");
                        outputBuilder.AppendLine($"[CL_LO_SRC]\n");

                        // ---------------------------LO_Frequency -------------------------------
                        int numCols = 0;
                        if (int.TryParse(worksheet.Cells["B1"].Value?.ToString(), out int n_value))
                        {
                            numCols = n_value;
                        }

                        int numRows = 0;
                        if (int.TryParse(worksheet.Cells["A2"].Value?.ToString(), out int m_value))
                        {
                            numRows = m_value;
                        }

                        if (numCols <= 0 || numRows <= 0)
                        {
                            Console.WriteLine("Value unvalid.");
                            outputBuilder.AppendLine($"Value unvalid.");
                            File.WriteAllText(outputMixer, outputBuilder.ToString());
<<<<<<< HEAD
                            MessageBox.Show("Values at B1 and A2 are unvalid or null. Model Mixer can not be converted", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
=======
>>>>>>> 2742ae2f9991af8eb6b861e2a4bb4cf034f1f554
                            return;
                        }
                        int startRow = 4;
                        int startCol = 2;
                        int counter = 1;

                        double ParseScientificNumber(string sciNumber)
                        {
                            if (string.IsNullOrEmpty(sciNumber))
                                return 0;

                            // Xử lý chuỗi dạng "1410e6"
                            if (sciNumber.Contains("e") || sciNumber.Contains("E"))
                            {
                                // Tách phần cơ số và số mũ
                                char[] separators = new char[] { 'e', 'E' };
                                string[] parts = sciNumber.Split(separators, 2);

                                if (parts.Length == 2)
                                {
                                    if (double.TryParse(parts[0], out double baseValue) &&
                                        double.TryParse(parts[1], out double exponent))
                                    {
                                        return baseValue * Math.Pow(10, exponent);
                                    }
                                }
                            }

                            // Xử lý chuỗi số thông thường
                            if (double.TryParse(sciNumber, out double result))
                                return result;

                            // Trả về 0 nếu không thể chuyển đổi
                            return 0;
                        }
                        string FormatNumber(string number)
                        {
                            if (string.IsNullOrEmpty(number))
                                return "N/A";

                            // Kiểm tra nếu chuỗi kết thúc bằng 6 số 0
                            if (number.EndsWith("000000"))
                            {
                                return number.Substring(0, number.Length - 6) + "e6";
                            }
                            return number;
                        }

                        // 3. Lặp qua từng ô trong ma trận đã xác định
                        for (int row = startRow; row < startRow + numRows; row++)
                        {
                            for (int col = startCol; col < startCol + numCols; col++)
                            {
                                // Lấy giá trị của ô hiện tại
                                var cellValue = worksheet.Cells[row, col].Value;
                                // 4. Kiểm tra xem ô có dữ liệu hay không (không rỗng, không null)
                                if (cellValue != null && !string.IsNullOrWhiteSpace(cellValue.ToString()))
                                {
                                    // Lấy giá trị string input từ cột A cùng hàng 
                                    var Input_LO = worksheet.Cells[row, 1].Value?.ToString();
                                    // Lấy giá trị string output từ hàng B cùng cột 
                                    var Output_LO = worksheet.Cells[3, col].Value?.ToString();
                                    // đổi sang double để tính toán 
                                    double inputValue = ParseScientificNumber(Input_LO);
                                    double outputValue = ParseScientificNumber(Output_LO);
                                    double Value_LO = inputValue - outputValue;
                                    //trả lại string 
                                    string formatted_LO = FormatNumber(Value_LO.ToString());
                                    //kết quả terminal 
                                    Console.WriteLine($"Freq{counter}={formatted_LO}");
                                    outputBuilder.AppendLine($"Freq{counter}={formatted_LO}");
                                    counter++;
                                }
                            }
                        }
                    }    // LO Frequency    [CL_LO_SRC]
                    {
                        //------------------POWER DRIVE ------------------------------
                        int numPowerInput = 0;
                        if (int.TryParse(worksheet.Cells["C2"].Value?.ToString(), out int p_value))
                        {
                            numPowerInput = p_value;
                        }

                        int numCols = 0;
                        if (int.TryParse(worksheet.Cells["B1"].Value?.ToString(), out int n_value))
                        {
                            numCols = n_value;
                        }

                        int numRows = 0;
                        if (int.TryParse(worksheet.Cells["A2"].Value?.ToString(), out int m_value))
                        {
                            numRows = m_value;
                        }

                        if (numCols <= 0 || numRows <= 0)
                        {
                            Console.WriteLine("Value unvalid.");
                            outputBuilder.AppendLine($"Value unvalid.");
                            File.WriteAllText(outputMixer, outputBuilder.ToString());
<<<<<<< HEAD
                            MessageBox.Show("Values at B1 and A2 are unvalid or null. Model Mixer can not be converted", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
=======
>>>>>>> 2742ae2f9991af8eb6b861e2a4bb4cf034f1f554
                            return;
                        }

                        int startRow = 4;
                        int startCol = 2;
                        int counter = 1;
                        for (int row = startRow; row < startRow + numRows; row++)
                        {
                            for (int col = startCol; col < startCol + numCols; col++)
                            {
                                // Lấy giá trị của ô hiện tại
                                var cellValue = worksheet.Cells[row, col].Value;
                                // 4. Kiểm tra xem ô có dữ liệu hay không (không rỗng, không null)
                                if (cellValue != null && !string.IsNullOrWhiteSpace(cellValue.ToString()))
                                {

                                    // Lấy giá trị input từ cột A cùng hàng
                                    var inputValue = worksheet.Cells[row, 1].Value?.ToString();
                                    // Xử lý định dạng số (thay thế 6 số 0 bằng 'e6')
                                    //string formattedInput = FormatNumber(inputValue);
                                    //kết quả terminal 
                                    Console.WriteLine($"Power{counter}={numPowerInput}");
                                    outputBuilder.AppendLine($"Power{counter}={numPowerInput}");
                                    counter++;
                                }
                            }
                        }

                    }    // POWER DRIVE           
                    {
                        //------------------SETPOWER DRIVE ------------------------------
                        int numCols = 0;
                        if (int.TryParse(worksheet.Cells["B1"].Value?.ToString(), out int n_value))
                        {
                            numCols = n_value;
                        }

                        int numRows = 0;
                        if (int.TryParse(worksheet.Cells["A2"].Value?.ToString(), out int m_value))
                        {
                            numRows = m_value;
                        }

                        if (numCols <= 0 || numRows <= 0)
                        {
                            Console.WriteLine("Value unvalid.");
                            outputBuilder.AppendLine($"Value unvalid");
                            File.WriteAllText(outputMixer, outputBuilder.ToString());
<<<<<<< HEAD
                            MessageBox.Show("Values at B1 and A2 are unvalid or null. Model Mixer can not be converted", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
=======
>>>>>>> 2742ae2f9991af8eb6b861e2a4bb4cf034f1f554
                            return;
                        }

                        int startRow = 4;
                        int startCol = 2;
                        int counter = 1;
                        for (int row = startRow; row < startRow + numRows; row++)
                        {
                            for (int col = startCol; col < startCol + numCols; col++)
                            {
                                // Lấy giá trị của ô hiện tại
                                var cellValue = worksheet.Cells[row, col].Value;
                                // 4. Kiểm tra xem ô có dữ liệu hay không (không rỗng, không null)
                                if (cellValue != null && !string.IsNullOrWhiteSpace(cellValue.ToString()))
                                {

                                    // Lấy giá trị input từ cột A cùng hàng
                                    var inputValue = worksheet.Cells[row, 1].Value?.ToString();
                                    // Xử lý định dạng số (thay thế 6 số 0 bằng 'e6')
                                    //string formattedInput = FormatNumber(inputValue);
                                    //kết quả terminal 
                                    Console.WriteLine($"SetPower{counter}=0");
                                    outputBuilder.AppendLine($"SetPower{counter}=0");
                                    counter++;
                                }
                            }
                        }
                    }    // SETPOWER DRIVE 
                    {
                        Console.WriteLine($"\n");
                        Console.WriteLine($"[CL_IF]");
                        outputBuilder.AppendLine($"\n");
                        outputBuilder.AppendLine($"[CL_IF]");

                        // ---------------------- OUTPUT FREQUENCY ------------------------------------
                        int numCols = 0;
                        if (int.TryParse(worksheet.Cells["B1"].Value?.ToString(), out int n_value))
                        {
                            numCols = n_value;
                        }

                        int numRows = 0;
                        if (int.TryParse(worksheet.Cells["A2"].Value?.ToString(), out int m_value))
                        {
                            numRows = m_value;
                        }

                        if (numCols <= 0 || numRows <= 0)
                        {
                            Console.WriteLine("Value unvalid.");
                            outputBuilder.AppendLine($"Value unvalid.");
                            File.WriteAllText(outputMixer, outputBuilder.ToString());
<<<<<<< HEAD
                            MessageBox.Show("Values at B1 and A2 are unvalid or null. Model Mixer can not be converted", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
=======
>>>>>>> 2742ae2f9991af8eb6b861e2a4bb4cf034f1f554
                            return;
                        }

                        int startRow = 4;
                        int startCol = 2;
                        int counter = 1;
                        // Hàm xử lý định dạng số (thay thế 6 số 0 cuối bằng 'e6')
                        string FormatNumber(string number)
                        {
                            if (string.IsNullOrEmpty(number))
                                return "N/A";

                            // Kiểm tra nếu chuỗi kết thúc bằng 6 số 0
                            if (number.EndsWith("000000"))
                            {
                                return number.Substring(0, number.Length - 6) + "e6";
                            }
                            return number;
                        }
                        // 3. Lặp qua từng ô trong ma trận đã xác định
                        for (int row = startRow; row < startRow + numRows; row++)
                        {
                            for (int col = startCol; col < startCol + numCols; col++)
                            {
                                // Lấy giá trị của ô hiện tại
                                var cellValue = worksheet.Cells[row, col].Value;
                                // 4. Kiểm tra xem ô có dữ liệu hay không (không rỗng, không null)
                                if (cellValue != null && !string.IsNullOrWhiteSpace(cellValue.ToString()))
                                {
                                    // Lấy giá trị input từ hàng B cùng cột
                                    var inputValue = worksheet.Cells[3, col].Value?.ToString();
                                    // Xử lý định dạng số (thay thế 6 số 0 bằng 'e6')
                                    string formattedInput = FormatNumber(inputValue);
                                    //kết quả terminal 
                                    Console.WriteLine($"Freq{counter}={formattedInput}");
                                    outputBuilder.AppendLine($"Freq{counter}={formattedInput}");
                                    counter++;
                                }
                            }
                        }
                    }    // OUTPUT FREQUENCY [CL_IF]
                    {
                        //------------------SPAN -----------------------------
                        int numPowerInput = 0;
                        if (int.TryParse(worksheet.Cells["E2"].Value?.ToString(), out int p_value))
                        {
                            numPowerInput = p_value;
                        }

                        int numCols = 0;
                        if (int.TryParse(worksheet.Cells["B1"].Value?.ToString(), out int n_value))
                        {
                            numCols = n_value;
                        }

                        int numRows = 0;
                        if (int.TryParse(worksheet.Cells["A2"].Value?.ToString(), out int m_value))
                        {
                            numRows = m_value;
                        }

                        if (numCols <= 0 || numRows <= 0)
                        {
                            Console.WriteLine("Value unvalid.");
                            outputBuilder.AppendLine($"Value unvalid.");
                            File.WriteAllText(outputMixer, outputBuilder.ToString());
<<<<<<< HEAD
                            MessageBox.Show("Values at B1 and A2 are unvalid or null. Model Mixer can not be converted", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
=======
>>>>>>> 2742ae2f9991af8eb6b861e2a4bb4cf034f1f554
                            return;
                        }
                        // Hàm xử lý định dạng số (thay thế 4 số 0 cuối bằng 'e4')
                        string FormatNumber(string number)
                        {
                            if (string.IsNullOrEmpty(number))
                                return "N/A";

                            // Kiểm tra nếu chuỗi kết thúc bằng 6 số 0
                            if (number.EndsWith("0000"))
                            {
                                return number.Substring(0, number.Length - 4) + "e4";
                            }
                            return number;
                        }
                        int startRow = 4;
                        int startCol = 2;
                        int counter = 1;
                        for (int row = startRow; row < startRow + numRows; row++)
                        {
                            for (int col = startCol; col < startCol + numCols; col++)
                            {
                                // Lấy giá trị của ô hiện tại
                                var cellValue = worksheet.Cells[row, col].Value;
                                // 4. Kiểm tra xem ô có dữ liệu hay không (không rỗng, không null)
                                if (cellValue != null && !string.IsNullOrWhiteSpace(cellValue.ToString()))
                                {

                                    // Lấy giá trị input từ cột A cùng hàng
                                    //var inputValue = worksheet.Cells[row, 1].Value?.ToString();
                                    //chuyển từ int sang string 
                                    string A = FormatNumber(numPowerInput.ToString());
                                    // Xử lý định dạng số (thay thế 4 số 0 bằng 'e4')
                                    string formattedInput = FormatNumber(A);
                                    //kết quả terminal 
                                    Console.WriteLine($"Span{counter}={formattedInput}");
                                    outputBuilder.AppendLine($"Span{counter}={formattedInput}");
                                    counter++;
                                }
                            }
                        }
                    }    // SPAN [CL_IF]
                    {
                        //------------------OFFSET -----------------------------
                        int numCols = 0;
                        if (int.TryParse(worksheet.Cells["B1"].Value?.ToString(), out int n_value))
                        {
                            numCols = n_value;
                        }

                        int numRows = 0;
                        if (int.TryParse(worksheet.Cells["A2"].Value?.ToString(), out int m_value))
                        {
                            numRows = m_value;
                        }

                        if (numCols <= 0 || numRows <= 0)
                        {
                            Console.WriteLine("Value unvalid.");
                            outputBuilder.AppendLine($"Value unvalid");
                            File.WriteAllText(outputMixer, outputBuilder.ToString());
<<<<<<< HEAD
                            MessageBox.Show("Values at B1 and A2 are unvalid or null. Model Mixer can not be converted", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
=======
>>>>>>> 2742ae2f9991af8eb6b861e2a4bb4cf034f1f554
                            return;
                        }

                        int startRow = 4;
                        int startCol = 2;
                        int counter = 1;
                        for (int row = startRow; row < startRow + numRows; row++)
                        {
                            for (int col = startCol; col < startCol + numCols; col++)
                            {
                                // Lấy giá trị của ô hiện tại
                                var cellValue = worksheet.Cells[row, col].Value;
                                // 4. Kiểm tra xem ô có dữ liệu hay không (không rỗng, không null)
                                if (cellValue != null && !string.IsNullOrWhiteSpace(cellValue.ToString()))
                                {

                                    // Lấy giá trị input từ cột A cùng hàng
                                    var inputValue = worksheet.Cells[row, 1].Value?.ToString();
                                    // Xử lý định dạng số (thay thế 6 số 0 bằng 'e6')
                                    //string formattedInput = FormatNumber(inputValue);
                                    //kết quả terminal 
                                    Console.WriteLine($"Offset{counter}=0");
                                    outputBuilder.AppendLine($"Offset{counter}=0");
                                    counter++;
                                }
                            }
                        }
                    }    // OFFSET [CL_IF]
                    {
                        //------------------SETPOWER CL_IF -----------------------------
                        int numCols = 0;
                        if (int.TryParse(worksheet.Cells["B1"].Value?.ToString(), out int n_value))
                        {
                            numCols = n_value;
                        }

                        int numRows = 0;
                        if (int.TryParse(worksheet.Cells["A2"].Value?.ToString(), out int m_value))
                        {
                            numRows = m_value;
                        }

                        if (numCols <= 0 || numRows <= 0)
                        {
                            Console.WriteLine("Value unvalid.");
                            outputBuilder.AppendLine($"Value unvalid.");
                            File.WriteAllText(outputMixer, outputBuilder.ToString());
<<<<<<< HEAD
                            MessageBox.Show("Values at B1 and A2 are unvalid or null. Model Mixer can not be converted", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
=======
>>>>>>> 2742ae2f9991af8eb6b861e2a4bb4cf034f1f554
                            return;
                        }

                        int startRow = 4;
                        int startCol = 2;
                        int counter = 1;
                        for (int row = startRow; row < startRow + numRows; row++)
                        {
                            for (int col = startCol; col < startCol + numCols; col++)
                            {
                                // Lấy giá trị của ô hiện tại
                                var cellValue = worksheet.Cells[row, col].Value;
                                // 4. Kiểm tra xem ô có dữ liệu hay không (không rỗng, không null)
                                if (cellValue != null && !string.IsNullOrWhiteSpace(cellValue.ToString()))
                                {

                                    // Lấy giá trị input từ cột A cùng hàng
                                    var inputValue = worksheet.Cells[row, 1].Value?.ToString();
                                    // Xử lý định dạng số (thay thế 6 số 0 bằng 'e6')
                                    //string formattedInput = FormatNumber(inputValue);
                                    //kết quả terminal 
                                    Console.WriteLine($"SetPower{counter}=0 ");
                                    outputBuilder.AppendLine($"SetPower{counter}=0");
                                    counter++;
                                }
                            }
                        }
                    }    // SETPOWER [CL_IF]
                    {
                        Console.WriteLine($"\n");
                        Console.WriteLine($"[CL_SPEC]\n");
                        outputBuilder.AppendLine($"\n");
                        outputBuilder.AppendLine($"[CL_SPEC]\n");

                        //------------------[CL_SPEC] -----------------------------
                        int numCols = 0;
                        if (int.TryParse(worksheet.Cells["B1"].Value?.ToString(), out int n_value))
                        {
                            numCols = n_value;
                        }

                        int numRows = 0;
                        if (int.TryParse(worksheet.Cells["A2"].Value?.ToString(), out int m_value))
                        {
                            numRows = m_value;
                        }

                        if (numCols <= 0 || numRows <= 0)
                        {
                            Console.WriteLine("Value unvalid.");
                            outputBuilder.AppendLine($"Value unvalid.");
                            File.WriteAllText(outputMixer, outputBuilder.ToString());
<<<<<<< HEAD
                            MessageBox.Show("Values at B1 and A2 are unvalid or null. Model Mixer can not be converted", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
=======
>>>>>>> 2742ae2f9991af8eb6b861e2a4bb4cf034f1f554
                            return;
                        }


                        double ParseScientificNumber(string sciNumber)
                        {
                            if (string.IsNullOrEmpty(sciNumber))
                                return 0;

                            // Xử lý chuỗi dạng "1410e6"
                            if (sciNumber.Contains("e") || sciNumber.Contains("E"))
                            {
                                // Tách phần cơ số và số mũ
                                char[] separators = new char[] { 'e', 'E' };
                                string[] parts = sciNumber.Split(separators, 2);

                                if (parts.Length == 2)
                                {
                                    if (double.TryParse(parts[0], out double baseValue) &&
                                        double.TryParse(parts[1], out double exponent))
                                    {
                                        return baseValue * Math.Pow(10, exponent);
                                    }
                                }
                            }

                            // Xử lý chuỗi số thông thường
                            if (double.TryParse(sciNumber, out double result))
                                return result;

                            // Trả về 0 nếu không thể chuyển đổi
                            return 0;
                        }
                        // Hàm xử lý định dạng số (thay thế 4 số 0 cuối bằng 'e4')
                        string FormatNumber(string number)
                        {
                            if (string.IsNullOrEmpty(number))
                                return "N/A";

                            // Kiểm tra nếu chuỗi kết thúc bằng 6 số 0
                            if (number.EndsWith("0000"))
                            {
                                return number.Substring(0, number.Length - 4) + "e4";
                            }
                            return number;
                        }

                        int startRow = 4;
                        int startCol = 2;
                        int counter = 1;
                        for (int row = startRow; row < startRow + numRows; row++)
                        {
                            for (int col = startCol; col < startCol + numCols; col++)
                            {
                                // Lấy giá trị của ô hiện tại
                                var cellValue = worksheet.Cells[row, col].Value;
                                // 4. Kiểm tra xem ô có dữ liệu hay không (không rỗng, không null)
                                if (cellValue != null && !string.IsNullOrWhiteSpace(cellValue.ToString()))
                                {
                                    //xử lý gíá trị a/b 
                                    if (cellValue.ToString().Contains("/"))
                                    {
                                        string[] parts = cellValue.ToString().Split('/');
                                        if (parts.Length == 2)
                                        {
                                            double part1 = ParseScientificNumber(parts[0]);
                                            double part2 = ParseScientificNumber(parts[1]);
                                            //lấy giá trị âm 
                                            double neg_part2 = -part2;
                                            //chuyển từ double sang string 
                                            string part_MIN = FormatNumber(neg_part2.ToString());

                                            //kết quả terminal 
                                            Console.WriteLine($"SpecMin{counter}= {part_MIN}");
                                            outputBuilder.AppendLine($"SpecMin{counter}= {part_MIN}");
                                            counter++;
                                        }
                                    }
                                    else
                                    {
                                        Console.WriteLine($"SpecMin{counter}= unvalid");
                                        outputBuilder.AppendLine($"SpecMin{counter}= unvalid");
                                        counter++;
                                    }

                                }

                            }
                        }
                    }    // [CL_SPEC_MIN]
                    {
                        //------------------[CL_SPEC_MAX] -----------------------------
                        int numCols = 0;
                        if (int.TryParse(worksheet.Cells["B1"].Value?.ToString(), out int n_value))
                        {
                            numCols = n_value;
                        }

                        int numRows = 0;
                        if (int.TryParse(worksheet.Cells["A2"].Value?.ToString(), out int m_value))
                        {
                            numRows = m_value;
                        }

                        if (numCols <= 0 || numRows <= 0)
                        {
                            Console.WriteLine("Value unvalid.");
                            outputBuilder.AppendLine($"Value unvalid.");
                            File.WriteAllText(outputMixer, outputBuilder.ToString());
<<<<<<< HEAD
                            MessageBox.Show("Values at B1 and A2 are unvalid or null. Model Mixer can not be converted", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
=======
>>>>>>> 2742ae2f9991af8eb6b861e2a4bb4cf034f1f554
                            return;
                        }


                        double ParseScientificNumber(string sciNumber)
                        {
                            if (string.IsNullOrEmpty(sciNumber))
                                return 0;

                            // Xử lý chuỗi dạng "1410e6"
                            if (sciNumber.Contains("e") || sciNumber.Contains("E"))
                            {
                                // Tách phần cơ số và số mũ
                                char[] separators = new char[] { 'e', 'E' };
                                string[] parts = sciNumber.Split(separators, 2);

                                if (parts.Length == 2)
                                {
                                    if (double.TryParse(parts[0], out double baseValue) &&
                                        double.TryParse(parts[1], out double exponent))
                                    {
                                        return baseValue * Math.Pow(10, exponent);
                                    }
                                }
                            }

                            // Xử lý chuỗi số thông thường
                            if (double.TryParse(sciNumber, out double result))
                                return result;

                            // Trả về 0 nếu không thể chuyển đổi
                            return 0;
                        }
                        // Hàm xử lý định dạng số (thay thế 4 số 0 cuối bằng 'e4')
                        string FormatNumber(string number)
                        {
                            if (string.IsNullOrEmpty(number))
                                return "N/A";

                            // Kiểm tra nếu chuỗi kết thúc bằng 6 số 0
                            if (number.EndsWith("0000"))
                            {
                                return number.Substring(0, number.Length - 4) + "e4";
                            }
                            return number;
                        }

                        int startRow = 4;
                        int startCol = 2;
                        int counter = 1;
                        for (int row = startRow; row < startRow + numRows; row++)
                        {
                            for (int col = startCol; col < startCol + numCols; col++)
                            {
                                // Lấy giá trị của ô hiện tại
                                var cellValue = worksheet.Cells[row, col].Value;
                                // 4. Kiểm tra xem ô có dữ liệu hay không (không rỗng, không null)
                                if (cellValue != null && !string.IsNullOrWhiteSpace(cellValue.ToString()))
                                {
                                    //xử lý gíá trị a/b 
                                    if (cellValue.ToString().Contains("/"))
                                    {
                                        string[] parts = cellValue.ToString().Split('/');
                                        if (parts.Length == 2)
                                        {
                                            double part1 = ParseScientificNumber(parts[0]);
                                            double part2 = ParseScientificNumber(parts[1]);
                                            //lấy giá trị âm 
                                            double neg_part1 = -part1;
                                            //chuyển từ double sang string 
                                            string part_MAX = FormatNumber(neg_part1.ToString());

                                            //kết quả terminal 
                                            Console.WriteLine($"SpecMax{counter}= {part_MAX}");
                                            outputBuilder.AppendLine($"SpecMax{counter}= {part_MAX}");
                                            counter++;

                                        }
                                    }
                                    else
                                    {
                                        Console.WriteLine($"SpecMax{counter}= unvalid");
                                        outputBuilder.AppendLine($"SpecMax{counter}= unvalid");
                                        counter++;
                                    }

                                }

                            }
                        }
                    }    // [CL_SPEC_MAX]
                    {
                        Console.WriteLine($"\n");
                        Console.WriteLine($"[TOI_RF_SRC1]");
                        outputBuilder.AppendLine($"\n");
                        outputBuilder.AppendLine($"[TOI_RF_SCR1]");
                        //------------------[TOI_RF_SRC1] -----------------------------

                        // //   -----------------------------------------FREQ SC1------------------------------------------------ 
                        int numRows1 = 0;
                        if (int.TryParse(worksheet.Cells["A21"].Value?.ToString(), out int n_value))
                        {
                            numRows1 = n_value;
                        }
                        if (numRows1 <= 0)
                        {
                            Console.WriteLine("Value unvalid.");
                            outputBuilder.AppendLine($"Value unvalid.");
                            File.WriteAllText(outputMixer, outputBuilder.ToString());
<<<<<<< HEAD
                            MessageBox.Show("Value at A21 is unvalid or null. Model Mixer can not be converted", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
=======
>>>>>>> 2742ae2f9991af8eb6b861e2a4bb4cf034f1f554
                            return;
                        }
                        int startRow1 = 22;
                        int counter1 = 1;

                        // Hàm xử lý định dạng số (thay thế 6 số 0 cuối bằng 'e6')
                        string FormatNumber(string number)
                        {
                            if (string.IsNullOrEmpty(number))
                                return "N/A";

                            // Kiểm tra nếu chuỗi kết thúc bằng 6 số 0
                            if (number.EndsWith("000000"))
                            {
                                return number.Substring(0, number.Length - 6) + "e6";
                            }
                            return number;
                        }


                        // 3. Lặp qua từng ô trong ma trận đã xác định
                        for (int row1 = startRow1; row1 < startRow1 + numRows1; row1++)
                        {
                            // Lấy giá trị của ô hiện tại
                            var cellValue1 = worksheet.Cells[row1, 1].Value;
                            // 4. Kiểm tra xem ô có dữ liệu hay không (không rỗng, không null)
                            if (cellValue1 != null && !string.IsNullOrWhiteSpace(cellValue1.ToString()))
                            {
                                // Lấy giá trị input từ cột A cùng hàng
                                var inputValue1 = worksheet.Cells[row1, 1].Value?.ToString();
                                // Xử lý định dạng số (thay thế 6 số 0 bằng 'e6')
                                string formattedInput = FormatNumber(inputValue1);
                                //kết quả terminal 
                                Console.WriteLine($"Freq{counter1}={formattedInput}");
                                outputBuilder.AppendLine($"Freq{counter1}={formattedInput}");
                                counter1++;
                            }

                        }
                        // -----------------------------------------POWER SC1----------------------------------------------------------- 
                        int numRows2 = 0;
                        if (int.TryParse(worksheet.Cells["A21"].Value?.ToString(), out int m_value))
                        {
                            numRows2 = m_value;
                        }
                        if (numRows2 <= 0)
                        {
                            Console.WriteLine("Value unvalid.");
                            outputBuilder.AppendLine($"Value unvalid.");
                            File.WriteAllText(outputMixer, outputBuilder.ToString());
<<<<<<< HEAD
                            MessageBox.Show("Value at A21 is unvalid or null. Model Mixer can not be converted", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
=======
>>>>>>> 2742ae2f9991af8eb6b861e2a4bb4cf034f1f554
                            return;
                        }
                        int startRow2 = 22;
                        int counter2 = 1;
                        double ParseScientificNumber(string sciNumber)
                        {
                            if (string.IsNullOrEmpty(sciNumber))
                                return 0;

                            // Xử lý chuỗi dạng "1410e6"
                            if (sciNumber.Contains("e") || sciNumber.Contains("E"))
                            {
                                // Tách phần cơ số và số mũ
                                char[] separators = new char[] { 'e', 'E' };
                                string[] parts = sciNumber.Split(separators, 2);

                                if (parts.Length == 2)
                                {
                                    if (double.TryParse(parts[0], out double baseValue) &&
                                        double.TryParse(parts[1], out double exponent))
                                    {
                                        return baseValue * Math.Pow(10, exponent);
                                    }
                                }
                            }

                            // Xử lý chuỗi số thông thường
                            if (double.TryParse(sciNumber, out double result))
                                return result;

                            // Trả về 0 nếu không thể chuyển đổi
                            return 0;
                        }


                        // 3. Lặp qua từng ô trong ma trận đã xác định
                        for (int row2 = startRow2; row2 < startRow2 + numRows2; row2++)
                        {
                            // Lấy giá trị của ô hiện tại
                            var cellValue2 = worksheet.Cells[row2, 4].Value;
                            // 4. Kiểm tra xem ô có dữ liệu hay không (không rỗng, không null)
                            if (cellValue2 != null && !string.IsNullOrWhiteSpace(cellValue2.ToString()))
                            {
                                // Lấy giá trị input từ cột D cùng hàng
                                var inputValue2 = worksheet.Cells[row2, 4].Value?.ToString();
                                // đổi sang double để tính toán 
                                double power = ParseScientificNumber(inputValue2);
                                double result = power - 3;
                                //trả lại string 
                                string formattedresult = FormatNumber(result.ToString());
                                //kết quả terminal 
                                Console.WriteLine($"Power{counter2}={formattedresult}");
                                outputBuilder.AppendLine($"Power{counter2}={formattedresult}");
                                counter2++;
                            }
                        }
                        // -----------------------------------------SETPOWER SC1-----------------------------------------------------------
                        int numRows3 = 0;
                        if (int.TryParse(worksheet.Cells["A21"].Value?.ToString(), out int p_value))
                        {
                            numRows3 = p_value;
                        }
                        if (numRows3 <= 0)
                        {
                            Console.WriteLine("Value unvalid.");
                            outputBuilder.AppendLine($"Value unvalid.");
                            File.WriteAllText(outputMixer, outputBuilder.ToString());
<<<<<<< HEAD
                            MessageBox.Show("Value at A21 is unvalid or null. Model Mixer can not be converted", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
=======
>>>>>>> 2742ae2f9991af8eb6b861e2a4bb4cf034f1f554
                            return;
                        }
                        int startRow3 = 22;
                        int counter3 = 1;
                        for (int row3 = startRow3; row3 < startRow3 + numRows3; row3++)
                        {
                            // Lấy giá trị của ô hiện tại
                            var cellValue3 = worksheet.Cells[row3, 1].Value;
                            // 4. Kiểm tra xem ô có dữ liệu hay không (không rỗng, không null)
                            if (cellValue3 != null && !string.IsNullOrWhiteSpace(cellValue3.ToString()))
                            {
                                Console.WriteLine($"SetPower{counter3}=0");
                                outputBuilder.AppendLine($"SetPower{counter3}=0 ");
                                counter3++;
                            }
                        }
                    }    // [TOI_RF_SRC1]  Third Order Intercept 
                    {
                        Console.WriteLine($"\n");
                        Console.WriteLine($"[TOI_RF_SRC2]");
                        outputBuilder.AppendLine($"\n");
                        outputBuilder.AppendLine($"[TOI_RF_SRC2]");
                        //------------------[TOI_RF_SRC2] -----------------------------

                        // //   -----------------------------------------FREQ SC2 ------------------------------------------------ 
                        int numRows1 = 0;
                        if (int.TryParse(worksheet.Cells["A21"].Value?.ToString(), out int n_value))
                        {
                            numRows1 = n_value;
                        }
                        if (numRows1 <= 0)
                        {
                            Console.WriteLine("Value unvalid.");
                            outputBuilder.AppendLine($"Value unvalid.");
                            File.WriteAllText(outputMixer, outputBuilder.ToString());
<<<<<<< HEAD
                            MessageBox.Show("Value at A21 is unvalid or null. Model Mixer can not be converted", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
=======
>>>>>>> 2742ae2f9991af8eb6b861e2a4bb4cf034f1f554
                            return;
                        }
                        int startRow1 = 22;
                        int counter1 = 1;

                        // Hàm xử lý định dạng số (thay thế 6 số 0 cuối bằng 'e6')
                        string FormatNumber(string number)
                        {
                            if (string.IsNullOrEmpty(number))
                                return "N/A";

                            // Kiểm tra nếu chuỗi kết thúc bằng 6 số 0
                            if (number.EndsWith("000000"))
                            {
                                return number.Substring(0, number.Length - 6) + "e6";
                            }
                            return number;
                        }


                        // 3. Lặp qua từng ô trong ma trận đã xác định
                        for (int row1 = startRow1; row1 < startRow1 + numRows1; row1++)
                        {
                            // Lấy giá trị của ô hiện tại
                            var cellValue1 = worksheet.Cells[row1, 2].Value;
                            // 4. Kiểm tra xem ô có dữ liệu hay không (không rỗng, không null)
                            if (cellValue1 != null && !string.IsNullOrWhiteSpace(cellValue1.ToString()))
                            {
                                // Lấy giá trị input từ cột B  cùng hàng
                                var inputValue1 = worksheet.Cells[row1, 2].Value?.ToString();
                                // Xử lý định dạng số (thay thế 6 số 0 bằng 'e6')
                                string formattedInput = FormatNumber(inputValue1);
                                //kết quả terminal 
                                Console.WriteLine($"Freq{counter1}={formattedInput}");
                                outputBuilder.AppendLine($"Freq{counter1}={formattedInput}");
                                counter1++;
                            }

                        }
                        // -----------------------------------------POWER SC2----------------------------------------------------------- 
                        int numRows2 = 0;
                        if (int.TryParse(worksheet.Cells["A21"].Value?.ToString(), out int m_value))
                        {
                            numRows2 = m_value;
                        }
                        if (numRows2 <= 0)
                        {
                            Console.WriteLine("Value unvalid.");
                            outputBuilder.AppendLine($"Value unvalid.");
                            File.WriteAllText(outputMixer, outputBuilder.ToString());
<<<<<<< HEAD
                            MessageBox.Show("Value at A21 is unvalid or null. Model Mixer can not be converted", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
=======
>>>>>>> 2742ae2f9991af8eb6b861e2a4bb4cf034f1f554
                            return;
                        }
                        int startRow2 = 22;
                        int counter2 = 1;
                        double ParseScientificNumber(string sciNumber)
                        {
                            if (string.IsNullOrEmpty(sciNumber))
                                return 0;

                            // Xử lý chuỗi dạng "1410e6"
                            if (sciNumber.Contains("e") || sciNumber.Contains("E"))
                            {
                                // Tách phần cơ số và số mũ
                                char[] separators = new char[] { 'e', 'E' };
                                string[] parts = sciNumber.Split(separators, 2);

                                if (parts.Length == 2)
                                {
                                    if (double.TryParse(parts[0], out double baseValue) &&
                                        double.TryParse(parts[1], out double exponent))
                                    {
                                        return baseValue * Math.Pow(10, exponent);
                                    }
                                }
                            }

                            // Xử lý chuỗi số thông thường
                            if (double.TryParse(sciNumber, out double result))
                                return result;

                            // Trả về 0 nếu không thể chuyển đổi
                            return 0;
                        }


                        // 3. Lặp qua từng ô trong ma trận đã xác định
                        for (int row2 = startRow2; row2 < startRow2 + numRows2; row2++)
                        {
                            // Lấy giá trị của ô hiện tại
                            var cellValue2 = worksheet.Cells[row2, 4].Value;
                            // 4. Kiểm tra xem ô có dữ liệu hay không (không rỗng, không null)
                            if (cellValue2 != null && !string.IsNullOrWhiteSpace(cellValue2.ToString()))
                            {
                                // Lấy giá trị input từ cột D cùng hàng
                                var inputValue2 = worksheet.Cells[row2, 4].Value?.ToString();
                                // đổi sang double để tính toán 
                                double power = ParseScientificNumber(inputValue2);
                                double result = power - 3;
                                //trả lại string 
                                string formattedresult = FormatNumber(result.ToString());
                                //kết quả terminal 
                                Console.WriteLine($"Power{counter2}={formattedresult}");
                                outputBuilder.AppendLine($"Power{counter2}={formattedresult}");
                                counter2++;
                            }
                        }
                        // -----------------------------------------SETPOWER SC2-----------------------------------------------------------
                        int numRows3 = 0;
                        if (int.TryParse(worksheet.Cells["A21"].Value?.ToString(), out int p_value))
                        {
                            numRows3 = p_value;
                        }
                        if (numRows3 <= 0)
                        {
                            Console.WriteLine("Value unvalid.");
                            outputBuilder.AppendLine($"Value unvalid.");
                            File.WriteAllText(outputMixer, outputBuilder.ToString());
<<<<<<< HEAD
                            MessageBox.Show("Value at A21 is unvalid or null. Model Mixer can not be converted", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
=======
>>>>>>> 2742ae2f9991af8eb6b861e2a4bb4cf034f1f554
                            return;
                        }
                        int startRow3 = 22;
                        int counter3 = 1;
                        for (int row3 = startRow3; row3 < startRow3 + numRows3; row3++)
                        {
                            // Lấy giá trị của ô hiện tại
                            var cellValue3 = worksheet.Cells[row3, 2].Value;
                            // 4. Kiểm tra xem ô có dữ liệu hay không (không rỗng, không null)
                            if (cellValue3 != null && !string.IsNullOrWhiteSpace(cellValue3.ToString()))
                            {
                                Console.WriteLine($"SetPower{counter3}=0 ");
                                outputBuilder.AppendLine($"SetPower{counter3}=0 ");
                                counter3++;
                            }
                        }
                    }    // [TOI_RF_SRC2]  Third Order Intercept 
                    {
                        Console.WriteLine($"\n");
                        Console.WriteLine($"[TOI_LO_SRC]");
                        outputBuilder.AppendLine($"\n");
                        outputBuilder.AppendLine($"[TOI_LO_SRC]");
                        //------------------ [TOI_LO_SRC] -----------------------------

                        // //   ----------------------------------------[TOI_LO_SCR]------------------------------------------------ 
                        int numRows1 = 0;
                        if (int.TryParse(worksheet.Cells["A21"].Value?.ToString(), out int n_value))
                        {
                            numRows1 = n_value;
                        }
                        if (numRows1 <= 0)
                        {
                            Console.WriteLine("Value unvalid.");
                            outputBuilder.AppendLine($"Value unvalid.");
                            File.WriteAllText(outputMixer, outputBuilder.ToString());
<<<<<<< HEAD
                            MessageBox.Show("Value at A21 is unvalid or null. Model Mixer can not be converted", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
=======
>>>>>>> 2742ae2f9991af8eb6b861e2a4bb4cf034f1f554
                            return;
                        }
                        int startRow1 = 22;
                        int counter1 = 1;

                        // Hàm xử lý định dạng số (thay thế 6 số 0 cuối bằng 'e6')
                        string FormatNumber(string number)
                        {
                            if (string.IsNullOrEmpty(number))
                                return "N/A";

                            // Kiểm tra nếu chuỗi kết thúc bằng 6 số 0
                            if (number.EndsWith("000000"))
                            {
                                return number.Substring(0, number.Length - 6) + "e6";
                            }
                            return number;
                        }


                        // 3. Lặp qua từng ô trong ma trận đã xác định
                        for (int row1 = startRow1; row1 < startRow1 + numRows1; row1++)
                        {
                            // Lấy giá trị của ô hiện tại
                            var cellValue1 = worksheet.Cells[row1, 3].Value;
                            // 4. Kiểm tra xem ô có dữ liệu hay không (không rỗng, không null)
                            if (cellValue1 != null && !string.IsNullOrWhiteSpace(cellValue1.ToString()))
                            {
                                // Lấy giá trị input từ cột B  cùng hàng
                                var inputValue1 = worksheet.Cells[row1, 3].Value?.ToString();
                                // Xử lý định dạng số (thay thế 6 số 0 bằng 'e6')
                                string formattedInput = FormatNumber(inputValue1);
                                //kết quả terminal 
                                Console.WriteLine($"Freq{counter1}={formattedInput}");
                                outputBuilder.AppendLine($"Freq{counter1}={formattedInput}");
                                counter1++;
                            }
                        }
                        // -----------------------------------------POWER LO----------------------------------------------------------- 
                        int numRows2 = 0;
                        if (int.TryParse(worksheet.Cells["A21"].Value?.ToString(), out int m_value))
                        {
                            numRows2 = m_value;
                        }
                        if (numRows2 <= 0)
                        {
                            Console.WriteLine("Value unvalid.");
                            outputBuilder.AppendLine($"Value unvalid.");
                            File.WriteAllText(outputMixer, outputBuilder.ToString());
<<<<<<< HEAD
                            MessageBox.Show("Value at A21 is unvalid or null. Model Mixer can not be converted", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
=======
>>>>>>> 2742ae2f9991af8eb6b861e2a4bb4cf034f1f554
                            return;
                        }
                        int startRow2 = 22;
                        int counter2 = 1;
                        double ParseScientificNumber(string sciNumber)
                        {
                            if (string.IsNullOrEmpty(sciNumber))
                                return 0;

                            // Xử lý chuỗi dạng "1410e6"
                            if (sciNumber.Contains("e") || sciNumber.Contains("E"))
                            {
                                // Tách phần cơ số và số mũ
                                char[] separators = new char[] { 'e', 'E' };
                                string[] parts = sciNumber.Split(separators, 2);

                                if (parts.Length == 2)
                                {
                                    if (double.TryParse(parts[0], out double baseValue) &&
                                        double.TryParse(parts[1], out double exponent))
                                    {
                                        return baseValue * Math.Pow(10, exponent);
                                    }
                                }
                            }

                            // Xử lý chuỗi số thông thường
                            if (double.TryParse(sciNumber, out double result))
                                return result;

                            // Trả về 0 nếu không thể chuyển đổi
                            return 0;
                        }


                        // 3. Lặp qua từng ô trong ma trận đã xác định
                        for (int row2 = startRow2; row2 < startRow2 + numRows2; row2++)
                        {
                            // Lấy giá trị của ô hiện tại
                            var cellValue2 = worksheet.Cells[row2, 5].Value;
                            // 4. Kiểm tra xem ô có dữ liệu hay không (không rỗng, không null)
                            if (cellValue2 != null && !string.IsNullOrWhiteSpace(cellValue2.ToString()))
                            {
                                // Lấy giá trị input từ cột D cùng hàng
                                var inputValue2 = worksheet.Cells[row2, 5].Value?.ToString();
                                // đổi sang double để tính toán 
                                // double power = ParseScientificNumber(inputValue2);
                                // double result = power - 3;
                                //trả lại string 
                                string formattedresult = FormatNumber(inputValue2.ToString());
                                //kết quả terminal 
                                Console.WriteLine($"Power{counter2}={formattedresult}");
                                outputBuilder.AppendLine($"Power{counter2}={formattedresult}");
                                counter2++;
                            }
                        }
                        // -----------------------------------------SETPOWER SC2-----------------------------------------------------------
                        int numRows3 = 0;
                        if (int.TryParse(worksheet.Cells["A21"].Value?.ToString(), out int p_value))
                        {
                            numRows3 = p_value;
                        }
                        if (numRows3 <= 0)
                        {
                            Console.WriteLine("Value unvalid.");
                            outputBuilder.AppendLine($"Value unvalid.");
                            File.WriteAllText(outputMixer, outputBuilder.ToString());
<<<<<<< HEAD
                            MessageBox.Show("Value at A21 is unvalid or null. Model Mixer can not be converted", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
=======
>>>>>>> 2742ae2f9991af8eb6b861e2a4bb4cf034f1f554
                            return;
                        }
                        int startRow3 = 22;
                        int counter3 = 1;
                        for (int row3 = startRow3; row3 < startRow3 + numRows3; row3++)
                        {
                            // Lấy giá trị của ô hiện tại
                            var cellValue3 = worksheet.Cells[row3, 3].Value;
                            // 4. Kiểm tra xem ô có dữ liệu hay không (không rỗng, không null)
                            if (cellValue3 != null && !string.IsNullOrWhiteSpace(cellValue3.ToString()))
                            {
                                Console.WriteLine($"SetPower{counter3}=0 ");
                                outputBuilder.AppendLine($"SetPower{counter3}=0 ");
                                counter3++;
                            }
                        }
                    }    // [TOI_LO_SRC]   Third Order Intercept 
                    {
                        Console.WriteLine($"\n");
                        Console.WriteLine($"[TOI_IF1]");
                        outputBuilder.AppendLine($"\n");
                        outputBuilder.AppendLine($"[TOI_IF1]");
                        //------------------[TOI_IF1] -----------------------------

                        // //   -----------------------------------------FREQ IF1 ------------------------------------------------ 
                        int numRows1 = 0;
                        if (int.TryParse(worksheet.Cells["A21"].Value?.ToString(), out int n_value))
                        {
                            numRows1 = n_value;
                        }
                        if (numRows1 <= 0)
                        {
                            Console.WriteLine("Value unvalid.");
                            outputBuilder.AppendLine($"Value unvalid.");
                            File.WriteAllText(outputMixer, outputBuilder.ToString());
<<<<<<< HEAD
                            MessageBox.Show("Value at A21 is unvalid or null. Model Mixer can not be converted", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
=======
>>>>>>> 2742ae2f9991af8eb6b861e2a4bb4cf034f1f554
                            return;
                        }
                        int startRow1 = 22;
                        int counter1 = 1;
                        // Hàm xử lý định dạng số (thay thế 6 số 0 cuối bằng 'e6')
                        string FormatNumber(string number)
                        {
                            if (string.IsNullOrEmpty(number))
                                return "N/A";

                            // Kiểm tra nếu chuỗi kết thúc bằng 6 số 0
                            if (number.EndsWith("000000"))
                            {
                                return number.Substring(0, number.Length - 6) + "e6";
                            }
                            return number;
                        }

                        // 3. Lặp qua từng ô trong ma trận đã xác định
                        for (int row1 = startRow1; row1 < startRow1 + numRows1; row1++)
                        {
                            // Lấy giá trị của ô hiện tại
                            var cellValue1 = worksheet.Cells[row1, 6].Value;
                            // 4. Kiểm tra xem ô có dữ liệu hay không (không rỗng, không null)
                            if (cellValue1 != null && !string.IsNullOrWhiteSpace(cellValue1.ToString()))
                            {
                                // Lấy giá trị input từ cột B  cùng hàng
                                var inputValue1 = worksheet.Cells[row1, 6].Value?.ToString();
                                // Xử lý định dạng số (thay thế 6 số 0 bằng 'e6')
                                string formattedInput = FormatNumber(inputValue1);
                                //kết quả terminal 
                                Console.WriteLine($"Freq{counter1}={formattedInput}");
                                outputBuilder.AppendLine($"Freq{counter1}={formattedInput}");
                                counter1++;
                            }
                        }
                        // //   -----------------------------------------SPAN 1------------------------------------------------ 
                        int numRows2 = 0;
                        if (int.TryParse(worksheet.Cells["A21"].Value?.ToString(), out int m_value))
                        {
                            numRows2 = m_value;
                        }
                        if (numRows2 <= 0)
                        {
                            Console.WriteLine("Value unvalid.");
                            outputBuilder.AppendLine($"Value unvalid");
<<<<<<< HEAD
                            MessageBox.Show("Value at A21 is unvalid or null. Model Mixer can not be converted", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
=======
>>>>>>> 2742ae2f9991af8eb6b861e2a4bb4cf034f1f554
                            return;
                        }
                        int startRow2 = 22;
                        int counter2 = 1;
                        // Hàm xử lý định dạng số (thay thế 4 số 0 cuối bằng 'e4')
                        string FormatNumber1(string number)
                        {
                            if (string.IsNullOrEmpty(number))
                                return "N/A";

                            // Kiểm tra nếu chuỗi kết thúc bằng 6 số 0
                            if (number.EndsWith("0000"))
                            {
                                return number.Substring(0, number.Length - 4) + "e4";
                            }
                            return number;
                        }
                        // 3. Lặp qua từng ô trong ma trận đã xác định
                        for (int row2 = startRow2; row2 < startRow2 + numRows2; row2++)
                        {
                            // Lấy giá trị của ô hiện tại
                            var cellValue2 = worksheet.Cells[row2, 8].Value;
                            // 4. Kiểm tra xem ô có dữ liệu hay không (không rỗng, không null)
                            if (cellValue2 != null && !string.IsNullOrWhiteSpace(cellValue2.ToString()))
                            {
                                // Lấy giá trị input từ cột B  cùng hàng
                                var inputValue2 = worksheet.Cells[row2, 8].Value?.ToString();
                                // Xử lý định dạng số (thay thế 6 số 0 bằng 'e6')
                                string formattedInput2 = FormatNumber1(inputValue2);
                                //kết quả terminal 
                                Console.WriteLine($"Span{counter2}={formattedInput2}");
                                outputBuilder.AppendLine($"Span{counter2}={formattedInput2}");
                                counter2++;
                            }
                        }
                        // //  ---------------------------------------------OFFSET 1----------------------------------------------
                        int numRows3 = 0;
                        if (int.TryParse(worksheet.Cells["A21"].Value?.ToString(), out int p_value))
                        {
                            numRows3 = p_value;
                        }
                        if (numRows3 <= 0)
                        {
                            Console.WriteLine("Value unvalid.");
                            outputBuilder.AppendLine($"Value unvalid.");
                            File.WriteAllText(outputMixer, outputBuilder.ToString());
<<<<<<< HEAD
                            MessageBox.Show("Value at A21 is unvalid or null. Model Mixer can not be converted", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
=======
>>>>>>> 2742ae2f9991af8eb6b861e2a4bb4cf034f1f554
                            return;
                        }
                        int startRow3 = 22;
                        int counter3 = 1;
                        // 3. Lặp qua từng ô trong ma trận đã xác định
                        for (int row3 = startRow3; row3 < startRow3 + numRows3; row3++)
                        {
                            // Lấy giá trị của ô hiện tại
                            var cellValue3 = worksheet.Cells[row3, 1].Value;
                            if (cellValue3 != null && !string.IsNullOrWhiteSpace(cellValue3.ToString()))
                            {
                                Console.WriteLine($"Offset{counter3}=0");
                                outputBuilder.AppendLine($"Offset{counter3}=0 ");
                                counter3++;
                            }
                        }
                        // //  ----------------------------------------SETPOWER 1 ----------------------------------------------
                        int numRows4 = 0;
                        if (int.TryParse(worksheet.Cells["A21"].Value?.ToString(), out int o_value))
                        {
                            numRows4 = o_value;
                        }
                        if (numRows4 <= 0)
                        {
                            Console.WriteLine("Value unvalid.");
                            outputBuilder.AppendLine($"Value unvalid.");
                            File.WriteAllText(outputMixer, outputBuilder.ToString());
<<<<<<< HEAD
                            MessageBox.Show("Value at A21 is unvalid or null. Model Mixer can not be converted", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
=======
>>>>>>> 2742ae2f9991af8eb6b861e2a4bb4cf034f1f554
                            return;
                        }
                        int startRow4 = 22;
                        int counter4 = 1;
                        // 3. Lặp qua từng ô trong ma trận đã xác định
                        for (int row4 = startRow4; row4 < startRow4 + numRows4; row4++)
                        {
                            // Lấy giá trị của ô hiện tại
                            var cellValue4 = worksheet.Cells[row4, 1].Value;
                            if (cellValue4 != null && !string.IsNullOrWhiteSpace(cellValue4.ToString()))
                            {
                                Console.WriteLine($"SetPower{counter4}=0");
                                outputBuilder.AppendLine($"SetPower{counter4}=0 ");
                                counter4++;
                            }
                        }
                    }    // [TOI_IF1]      Third Order Intercept 
                    {
                        Console.WriteLine($"\n");
                        Console.WriteLine($"[TOI_IF2]");
                        outputBuilder.AppendLine($"\n");
                        outputBuilder.AppendLine($"[TOI_IF2]");
                        //------------------[TOI_IF2] -----------------------------

                        // //   -----------------------------------------FREQ IF2 ------------------------------------------------ 
                        int numRows1 = 0;
                        if (int.TryParse(worksheet.Cells["A21"].Value?.ToString(), out int n_value))
                        {
                            numRows1 = n_value;
                        }
                        if (numRows1 <= 0)
                        {
                            Console.WriteLine("Value unvalid.");
                            outputBuilder.AppendLine($"Value unvalid.");
                            File.WriteAllText(outputMixer, outputBuilder.ToString());
<<<<<<< HEAD
                            MessageBox.Show("Value at A21 is unvalid or null. Model Mixer can not be converted", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
=======
>>>>>>> 2742ae2f9991af8eb6b861e2a4bb4cf034f1f554
                            return;
                        }
                        int startRow1 = 22;
                        int counter1 = 1;
                        // Hàm xử lý định dạng số (thay thế 6 số 0 cuối bằng 'e6')
                        string FormatNumber(string number)
                        {
                            if (string.IsNullOrEmpty(number))
                                return "N/A";

                            // Kiểm tra nếu chuỗi kết thúc bằng 6 số 0
                            if (number.EndsWith("000000"))
                            {
                                return number.Substring(0, number.Length - 6) + "e6";
                            }
                            return number;
                        }

                        // 3. Lặp qua từng ô trong ma trận đã xác định
                        for (int row1 = startRow1; row1 < startRow1 + numRows1; row1++)
                        {
                            // Lấy giá trị của ô hiện tại
                            var cellValue1 = worksheet.Cells[row1, 7].Value;
                            // 4. Kiểm tra xem ô có dữ liệu hay không (không rỗng, không null)
                            if (cellValue1 != null && !string.IsNullOrWhiteSpace(cellValue1.ToString()))
                            {
                                // Lấy giá trị input từ cột B  cùng hàng
                                var inputValue1 = worksheet.Cells[row1, 7].Value?.ToString();
                                // Xử lý định dạng số (thay thế 6 số 0 bằng 'e6')
                                string formattedInput = FormatNumber(inputValue1);
                                //kết quả terminal 
                                Console.WriteLine($"Freq{counter1}={formattedInput}");
                                outputBuilder.AppendLine($"Freq{counter1}={formattedInput}");
                                counter1++;
                            }
                        }
                        // //   -----------------------------------------SPAN 2------------------------------------------------ 
                        int numRows2 = 0;
                        if (int.TryParse(worksheet.Cells["A21"].Value?.ToString(), out int m_value))
                        {
                            numRows2 = m_value;
                        }
                        if (numRows2 <= 0)
                        {
                            Console.WriteLine("Value unvalid.");
                            outputBuilder.AppendLine($"Value unvalid.");
                            File.WriteAllText(outputMixer, outputBuilder.ToString());
<<<<<<< HEAD
                            MessageBox.Show("Value at A21 is unvalid or null. Model Mixer can not be converted", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
=======
>>>>>>> 2742ae2f9991af8eb6b861e2a4bb4cf034f1f554
                            return;
                        }
                        int startRow2 = 22;
                        int counter2 = 1;
                        // Hàm xử lý định dạng số (thay thế 4 số 0 cuối bằng 'e4')
                        string FormatNumber1(string number)
                        {
                            if (string.IsNullOrEmpty(number))
                                return "N/A";

                            // Kiểm tra nếu chuỗi kết thúc bằng 6 số 0
                            if (number.EndsWith("0000"))
                            {
                                return number.Substring(0, number.Length - 4) + "e4";
                            }
                            return number;
                        }
                        // 3. Lặp qua từng ô trong ma trận đã xác định
                        for (int row2 = startRow2; row2 < startRow2 + numRows2; row2++)
                        {
                            // Lấy giá trị của ô hiện tại
                            var cellValue2 = worksheet.Cells[row2, 8].Value;
                            // 4. Kiểm tra xem ô có dữ liệu hay không (không rỗng, không null)
                            if (cellValue2 != null && !string.IsNullOrWhiteSpace(cellValue2.ToString()))
                            {
                                // Lấy giá trị input từ cột B  cùng hàng
                                var inputValue2 = worksheet.Cells[row2, 8].Value?.ToString();
                                // Xử lý định dạng số (thay thế 6 số 0 bằng 'e6')
                                string formattedInput2 = FormatNumber1(inputValue2);
                                //kết quả terminal 
                                Console.WriteLine($"Span{counter2}={formattedInput2}");
                                outputBuilder.AppendLine($"Span{counter2}={formattedInput2}");
                                counter2++;
                            }
                        }
                        // //  ---------------------------------------------OFFSET 2----------------------------------------------
                        int numRows3 = 0;
                        if (int.TryParse(worksheet.Cells["A21"].Value?.ToString(), out int p_value))
                        {
                            numRows3 = p_value;
                        }
                        if (numRows3 <= 0)
                        {
                            Console.WriteLine("Value unvalid.");
                            outputBuilder.AppendLine($"Value unvalid.");
                            File.WriteAllText(outputMixer, outputBuilder.ToString());
<<<<<<< HEAD
                            MessageBox.Show("Value at A21 is unvalid or null. Model Mixer can not be converted", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
=======
>>>>>>> 2742ae2f9991af8eb6b861e2a4bb4cf034f1f554
                            return;
                        }
                        int startRow3 = 22;
                        int counter3 = 1;
                        // 3. Lặp qua từng ô trong ma trận đã xác định
                        for (int row3 = startRow3; row3 < startRow3 + numRows3; row3++)
                        {
                            // Lấy giá trị của ô hiện tại
                            var cellValue3 = worksheet.Cells[row3, 1].Value;
                            if (cellValue3 != null && !string.IsNullOrWhiteSpace(cellValue3.ToString()))
                            {
                                Console.WriteLine($"Offset{counter3}=0");
                                outputBuilder.AppendLine($"Offset{counter3}=0 ");
                                counter3++;
                            }
                        }
                        // //  ----------------------------------------SETPOWER 2 ----------------------------------------------
                        int numRows4 = 0;
                        if (int.TryParse(worksheet.Cells["A21"].Value?.ToString(), out int o_value))
                        {
                            numRows4 = o_value;
                        }
                        if (numRows4 <= 0)
                        {
                            Console.WriteLine("Value unvalid.");
                            outputBuilder.AppendLine($"Value unvalid.");
                            File.WriteAllText(outputMixer, outputBuilder.ToString());
<<<<<<< HEAD
                            MessageBox.Show("Value at A21 is unvalid or null. Model Mixer can not be converted", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
=======
>>>>>>> 2742ae2f9991af8eb6b861e2a4bb4cf034f1f554
                            return;
                        }
                        int startRow4 = 22;
                        int counter4 = 1;
                        // 3. Lặp qua từng ô trong ma trận đã xác định
                        for (int row4 = startRow4; row4 < startRow4 + numRows4; row4++)
                        {
                            // Lấy giá trị của ô hiện tại
                            var cellValue4 = worksheet.Cells[row4, 1].Value;
                            if (cellValue4 != null && !string.IsNullOrWhiteSpace(cellValue4.ToString()))
                            {
                                Console.WriteLine($"SetPower{counter4}=0");
                                outputBuilder.AppendLine($"SetPower{counter4}=0");
                                counter4++;
                            }
                        }
                    }    // [TOI_IF2]     
                    {
                        Console.WriteLine($"\n");
                        Console.WriteLine($"[TOI_SPEC]");
                        outputBuilder.AppendLine($"\n");
                        outputBuilder.AppendLine($"[TOI_SPEC]");
                        //------------------[TOI_SPEC] -----------------------------

                        // //   -----------------------------------------SPECMIN ------------------------------------------------ 
                        int numRows1 = 0;
                        if (int.TryParse(worksheet.Cells["A21"].Value?.ToString(), out int n_value))
                        {
                            numRows1 = n_value;
                        }
                        if (numRows1 <= 0)
                        {
                            Console.WriteLine("Value unvalid.");
                            outputBuilder.AppendLine($"Value unvalid.");
                            File.WriteAllText(outputMixer, outputBuilder.ToString());
<<<<<<< HEAD
                            MessageBox.Show("Value at A21 is unvalid or null. Model Mixer can not be converted", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
=======
>>>>>>> 2742ae2f9991af8eb6b861e2a4bb4cf034f1f554
                            return;
                        }
                        int startRow1 = 22;
                        int counter1 = 1;

                        // 3. Lặp qua từng ô trong ma trận đã xác định
                        for (int row1 = startRow1; row1 < startRow1 + numRows1; row1++)
                        {
                            // Lấy giá trị của ô hiện tại
                            var cellValue1 = worksheet.Cells[row1, 9].Value;
                            // 4. Kiểm tra xem ô có dữ liệu hay không (không rỗng, không null)
                            if (cellValue1 != null && !string.IsNullOrWhiteSpace(cellValue1.ToString()))
                            {
                                // Lấy giá trị input từ cột B  cùng hàng
                                var inputValue1 = worksheet.Cells[row1, 9].Value?.ToString();
                                // Xử lý định dạng số (thay thế 6 số 0 bằng 'e6')
                                //string formattedInput = FormatNumber(inputValue1);
                                //kết quả terminal 
                                Console.WriteLine($"SpecMin{counter1}={inputValue1}");
                                outputBuilder.AppendLine($"SpecMin{counter1}={inputValue1}");
                                counter1++;
                            }
                        }
                        // //   -----------------------------------------SPECMAX ------------------------------------------------ 
                        int numRows2 = 0;
                        if (int.TryParse(worksheet.Cells["A21"].Value?.ToString(), out int m_value))
                        {
                            numRows2 = m_value;
                        }
                        if (numRows2 <= 0)
                        {
                            Console.WriteLine("Value unvalid.");
                            outputBuilder.AppendLine($"Value unvalid.");
                            File.WriteAllText(outputMixer, outputBuilder.ToString());
<<<<<<< HEAD
                            MessageBox.Show("Value at A21 is unvalid or null. Model Mixer can not be converted", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
=======
>>>>>>> 2742ae2f9991af8eb6b861e2a4bb4cf034f1f554
                            return;
                        }
                        int startRow2 = 22;
                        int counter2 = 1;

                        // 3. Lặp qua từng ô trong ma trận đã xác định
                        for (int row2 = startRow2; row2 < startRow2 + numRows2; row2++)
                        {
                            // Lấy giá trị của ô hiện tại
                            var cellValue2 = worksheet.Cells[row2, 9].Value;
                            // 4. Kiểm tra xem ô có dữ liệu hay không (không rỗng, không null)
                            if (cellValue2 != null && !string.IsNullOrWhiteSpace(cellValue2.ToString()))
                            {
                                Console.WriteLine($"SpecMax{counter2}=99");
                                outputBuilder.AppendLine($"SpecMax{counter2}=99");
                                counter2++;
                            }
                        }
                        // //   -----------------------------------------OFFSET ------------------------------------------------ 
                        int numRows3 = 0;
                        if (int.TryParse(worksheet.Cells["A21"].Value?.ToString(), out int o_value))
                        {
                            numRows3 = o_value;
                        }
                        if (numRows3 <= 0)
                        {
                            Console.WriteLine("Value unvalid.");
                            outputBuilder.AppendLine($"Value unvalid.");
                            File.WriteAllText(outputMixer, outputBuilder.ToString());
<<<<<<< HEAD
                            MessageBox.Show("Value at A21 is unvalid or null. Model Mixer can not be converted", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
=======
>>>>>>> 2742ae2f9991af8eb6b861e2a4bb4cf034f1f554
                            return;
                        }
                        int startRow3 = 22;
                        int counter3 = 1;

                        // 3. Lặp qua từng ô trong ma trận đã xác định
                        for (int row3 = startRow3; row3 < startRow3 + numRows3; row3++)
                        {
                            // Lấy giá trị của ô hiện tại
                            var cellValue3 = worksheet.Cells[row3, 9].Value;
                            // 4. Kiểm tra xem ô có dữ liệu hay không (không rỗng, không null)
                            if (cellValue3 != null && !string.IsNullOrWhiteSpace(cellValue3.ToString()))
                            {
                                Console.WriteLine($"Offset{counter3}=0 ");
                                outputBuilder.AppendLine($"Offset{counter3}=0 ");
                                counter3++;
                            }
                        }
                    }    // [TOI_SPEC]
                    {
                        Console.WriteLine($"\n");
                        Console.WriteLine($"[ISO_LO]");
                        outputBuilder.AppendLine($"\n");
                        outputBuilder.AppendLine($"[ISO_LO]");
                        //------------------[ISO_LO] -----------------------------

                        // //   ----------------------------------------- FREQ ------------------------------------------------ 
                        int numRows1 = 0;
                        if (int.TryParse(worksheet.Cells["A41"].Value?.ToString(), out int n_value))
                        {
                            numRows1 = n_value;
                        }
                        if (numRows1 <= 0)
                        {
                            Console.WriteLine("Value unvalid.");
                            outputBuilder.AppendLine($"Value unvalid.");
                            File.WriteAllText(outputMixer, outputBuilder.ToString());
<<<<<<< HEAD
                            MessageBox.Show("Value at A41 is unvalid or null. Model Mixer can not be converted", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
=======
>>>>>>> 2742ae2f9991af8eb6b861e2a4bb4cf034f1f554
                            return;
                        }
                        int startRow1 = 42;
                        int counter1 = 1;
                        // Hàm xử lý định dạng số (thay thế 6 số 0 cuối bằng 'e6')
                        string FormatNumber(string number)
                        {
                            if (string.IsNullOrEmpty(number))
                                return "N/A";

                            // Kiểm tra nếu chuỗi kết thúc bằng 6 số 0
                            if (number.EndsWith("000000"))
                            {
                                return number.Substring(0, number.Length - 6) + "e6";
                            }
                            return number;
                        }

                        // 3. Lặp qua từng ô trong ma trận đã xác định
                        for (int row1 = startRow1; row1 < startRow1 + numRows1; row1++)
                        {
                            // Lấy giá trị của ô hiện tại
                            var cellValue1 = worksheet.Cells[row1, 1].Value;
                            // 4. Kiểm tra xem ô có dữ liệu hay không (không rỗng, không null)
                            if (cellValue1 != null && !string.IsNullOrWhiteSpace(cellValue1.ToString()))
                            {
                                // Lấy giá trị input từ cột A cùng hàng
                                var inputValue = worksheet.Cells[row1, 1].Value?.ToString();
                                // Xử lý định dạng số (thay thế 6 số 0 bằng 'e6')
                                string formattedInput = FormatNumber(inputValue);
                                //kết quả terminal 
                                Console.WriteLine($"Freq{counter1}={formattedInput}");
                                outputBuilder.AppendLine($"Freq{counter1}={formattedInput}");
                                counter1++;
                            }
                        }
                        // //   ----------------------------------------- POWER ------------------------------------------------ 
                        int numRows2 = 0;
                        if (int.TryParse(worksheet.Cells["A41"].Value?.ToString(), out int m_value))
                        {
                            numRows2 = m_value;
                        }
                        if (numRows2 <= 0)
                        {
                            Console.WriteLine("Value unvalid.");
                            outputBuilder.AppendLine($"Value unvalid.");
                            File.WriteAllText(outputMixer, outputBuilder.ToString());
<<<<<<< HEAD
                            MessageBox.Show("Value at A41 is unvalid or null. Model Mixer can not be converted", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
=======
>>>>>>> 2742ae2f9991af8eb6b861e2a4bb4cf034f1f554
                            return;
                        }
                        int startRow2 = 42;
                        int counter2 = 1;


                        // 3. Lặp qua từng ô trong ma trận đã xác định
                        for (int row2 = startRow2; row2 < startRow2 + numRows2; row2++)
                        {
                            // Lấy giá trị của ô hiện tại
                            var cellValue2 = worksheet.Cells[row2, 1].Value;
                            // 4. Kiểm tra xem ô có dữ liệu hay không (không rỗng, không null)
                            if (cellValue2 != null && !string.IsNullOrWhiteSpace(cellValue2.ToString()))
                            {
                                // Lấy giá trị input từ cột A cùng hàng
                                var inputValue2 = worksheet.Cells[row2, 2].Value?.ToString();
                                // Xử lý định dạng số (thay thế 6 số 0 bằng 'e6')
                                //string formattedInput2  = FormatNumber(inputValue2);
                                //kết quả terminal 
                                Console.WriteLine($"Power{counter2}={inputValue2}");
                                outputBuilder.AppendLine($"Power{counter2}={inputValue2}");
                                counter2++;
                            }
                        }
                        // //   ----------------------------------------- SETPOWER ------------------------------------------------ 
                        int numRows3 = 0;
                        if (int.TryParse(worksheet.Cells["A41"].Value?.ToString(), out int p_value))
                        {
                            numRows3 = p_value;
                        }
                        if (numRows3 <= 0)
                        {
                            Console.WriteLine("Value unvalid.");
                            outputBuilder.AppendLine($"Value unvalid.");
                            File.WriteAllText(outputMixer, outputBuilder.ToString());
<<<<<<< HEAD
                            MessageBox.Show("Value at A41 is unvalid or null. Model Mixer can not be converted", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
=======
>>>>>>> 2742ae2f9991af8eb6b861e2a4bb4cf034f1f554
                            return;
                        }
                        int startRow3 = 42;
                        int counter3 = 1;


                        // 3. Lặp qua từng ô trong ma trận đã xác định
                        for (int row3 = startRow3; row3 < startRow3 + numRows3; row3++)
                        {
                            // Lấy giá trị của ô hiện tại
                            var cellValue3 = worksheet.Cells[row3, 1].Value;
                            // 4. Kiểm tra xem ô có dữ liệu hay không (không rỗng, không null)
                            if (cellValue3 != null && !string.IsNullOrWhiteSpace(cellValue3.ToString()))
                            {
                                // Lấy giá trị input từ cột A cùng hàng
                                //var inputValue2 = worksheet.Cells[row2, 2].Value?.ToString();
                                // Xử lý định dạng số (thay thế 6 số 0 bằng 'e6')
                                //string formattedInput2  = FormatNumber(inputValue2);
                                //kết quả terminal 
                                Console.WriteLine($"SetPower{counter3}=0");
                                outputBuilder.AppendLine($"SetPower{counter3}=0");
                                counter3++;
                            }
                        }
                    }    // [ISO_LO]          Isolation 
                    {
                        Console.WriteLine($"\n");
                        Console.WriteLine($"[ISO_IF]");
                        outputBuilder.AppendLine($"\n");
                        outputBuilder.AppendLine($"[ISO_IF]");
                        //------------------[ISO_IF] -----------------------------

                        // //   ----------------------------------------- FREQ ------------------------------------------------ 
                        int numRows1 = 0;
                        if (int.TryParse(worksheet.Cells["A41"].Value?.ToString(), out int n_value))
                        {
                            numRows1 = n_value;
                        }
                        if (numRows1 <= 0)
                        {
                            Console.WriteLine("Value unvalid.");
                            outputBuilder.AppendLine($"Value unvalid.");
                            File.WriteAllText(outputMixer, outputBuilder.ToString());
<<<<<<< HEAD
                            MessageBox.Show("Value at A41 is unvalid or null. Model Mixer can not be converted", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
=======
>>>>>>> 2742ae2f9991af8eb6b861e2a4bb4cf034f1f554
                            return;
                        }
                        int startRow1 = 42;
                        int counter1 = 1;
                        // Hàm xử lý định dạng số (thay thế 6 số 0 cuối bằng 'e6')
                        string FormatNumber(string number)
                        {
                            if (string.IsNullOrEmpty(number))
                                return "N/A";

                            // Kiểm tra nếu chuỗi kết thúc bằng 6 số 0
                            if (number.EndsWith("000000"))
                            {
                                return number.Substring(0, number.Length - 6) + "e6";
                            }
                            return number;
                        }

                        // 3. Lặp qua từng ô trong ma trận đã xác định
                        for (int row1 = startRow1; row1 < startRow1 + numRows1; row1++)
                        {
                            // Lấy giá trị của ô hiện tại
                            var cellValue1 = worksheet.Cells[row1, 1].Value;
                            // 4. Kiểm tra xem ô có dữ liệu hay không (không rỗng, không null)
                            if (cellValue1 != null && !string.IsNullOrWhiteSpace(cellValue1.ToString()))
                            {
                                // Lấy giá trị input từ cột A cùng hàng
                                var inputValue = worksheet.Cells[row1, 1].Value?.ToString();
                                // Xử lý định dạng số (thay thế 6 số 0 bằng 'e6')
                                string formattedInput = FormatNumber(inputValue);
                                //kết quả terminal 
                                Console.WriteLine($"Freq{counter1}={formattedInput}");
                                outputBuilder.AppendLine($"Freq{counter1}={formattedInput}");
                                counter1++;
                            }
                        }
                        // //   ----------------------------------------- SPAN ------------------------------------------------ 
                        int numRows2 = 0;
                        if (int.TryParse(worksheet.Cells["A41"].Value?.ToString(), out int m_value))
                        {
                            numRows2 = m_value;
                        }
                        if (numRows2 <= 0)
                        {
                            Console.WriteLine("Value unvalid.");
                            outputBuilder.AppendLine($"Value unvalid.");
                            File.WriteAllText(outputMixer, outputBuilder.ToString());
<<<<<<< HEAD
                            MessageBox.Show("Value at A41 is unvalid or null. Model Mixer can not be converted", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
=======
>>>>>>> 2742ae2f9991af8eb6b861e2a4bb4cf034f1f554
                            return;
                        }
                        int startRow2 = 42;
                        int counter2 = 1;
                        // 3. Lặp qua từng ô trong ma trận đã xác định
                        for (int row2 = startRow2; row2 < startRow2 + numRows2; row2++)
                        {
                            // Lấy giá trị của ô hiện tại
                            var cellValue2 = worksheet.Cells[row2, 1].Value;
                            // 4. Kiểm tra xem ô có dữ liệu hay không (không rỗng, không null)
                            if (cellValue2 != null && !string.IsNullOrWhiteSpace(cellValue2.ToString()))
                            {
                                // Lấy giá trị input từ cột A cùng hàng
                                var inputValue2 = worksheet.Cells[row2, 3].Value?.ToString();
                                // Xử lý định dạng số (thay thế 6 số 0 bằng 'e6')
                                string formattedInput2 = FormatNumber(inputValue2);
                                //kết quả terminal 
                                Console.WriteLine($"Span{counter2}={formattedInput2}");
                                outputBuilder.AppendLine($"Span{counter2}={formattedInput2}");
                                counter2++;
                            }
                        }
                        // //   ----------------------------------------- OFFSET ------------------------------------------------ 
                        int numRows3 = 0;
                        if (int.TryParse(worksheet.Cells["A41"].Value?.ToString(), out int p_value))
                        {
                            numRows3 = p_value;
                        }
                        if (numRows3 <= 0)
                        {
                            Console.WriteLine("Value unvalid.");
                            outputBuilder.AppendLine($"Value unvalid.");
                            File.WriteAllText(outputMixer, outputBuilder.ToString());
<<<<<<< HEAD
                            MessageBox.Show("Value at A41 is unvalid or null. Model Mixer can not be converted", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
=======
>>>>>>> 2742ae2f9991af8eb6b861e2a4bb4cf034f1f554
                            return;
                        }
                        int startRow3 = 42;
                        int counter3 = 1;
                        // 3. Lặp qua từng ô trong ma trận đã xác định
                        for (int row3 = startRow3; row3 < startRow3 + numRows3; row3++)
                        {
                            // Lấy giá trị của ô hiện tại
                            var cellValue3 = worksheet.Cells[row3, 1].Value;
                            // 4. Kiểm tra xem ô có dữ liệu hay không (không rỗng, không null)
                            if (cellValue3 != null && !string.IsNullOrWhiteSpace(cellValue3.ToString()))
                            {
                                // Lấy giá trị input từ cột A cùng hàng
                                //var inputValue3  = worksheet.Cells[row3 , 3].Value?.ToString();
                                // Xử lý định dạng số (thay thế 6 số 0 bằng 'e6')
                                //string formattedInput2 = FormatNumber(inputValue2);
                                //kết quả terminal 
                                Console.WriteLine($"Offset{counter3}=0");
                                outputBuilder.AppendLine($"Offset{counter3}=0");
                                counter3++;
                            }
                        }
                        // //   ----------------------------------------- SETPOWER ------------------------------------------------ 
                        int numRows4 = 0;
                        if (int.TryParse(worksheet.Cells["A41"].Value?.ToString(), out int o_value))
                        {
                            numRows4 = o_value;
                        }
                        if (numRows4 <= 0)
                        {
                            Console.WriteLine("Value unvalid.");
                            outputBuilder.AppendLine($"Value unvalid.");
                            File.WriteAllText(outputMixer, outputBuilder.ToString());
<<<<<<< HEAD
                            MessageBox.Show("Value at A41 is unvalid or null. Model Mixer can not be converted", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
=======
>>>>>>> 2742ae2f9991af8eb6b861e2a4bb4cf034f1f554
                            return;
                        }
                        int startRow4 = 42;
                        int counter4 = 1;
                        // 3. Lặp qua từng ô trong ma trận đã xác định
                        for (int row4 = startRow4; row4 < startRow4 + numRows4; row4++)
                        {
                            // Lấy giá trị của ô hiện tại
                            var cellValue4 = worksheet.Cells[row4, 1].Value;
                            // 4. Kiểm tra xem ô có dữ liệu hay không (không rỗng, không null)
                            if (cellValue4 != null && !string.IsNullOrWhiteSpace(cellValue4.ToString()))
                            {
                                // Lấy giá trị input từ cột A cùng hàng
                                //var inputValue3  = worksheet.Cells[row3 , 3].Value?.ToString();
                                // Xử lý định dạng số (thay thế 6 số 0 bằng 'e6')
                                //string formattedInput2 = FormatNumber(inputValue2);
                                //kết quả terminal 
                                Console.WriteLine($"SetPower{counter4}=0");
                                outputBuilder.AppendLine($"SetPower{counter4}=0");
                                counter4++;
                            }
                        }
                    }    // [ISO_IF] 
                    {
                        Console.WriteLine($"\n");
                        Console.WriteLine($"[ISO_SPEC]");
                        outputBuilder.AppendLine($"\n");
                        outputBuilder.AppendLine($"[ISO_SPEC]");
                        //------------------[ISO_SPEC] -----------------------------

                        // //   ----------------------------------------- SPECMIN------------------------------------------------ 
                        int numRows1 = 0;
                        if (int.TryParse(worksheet.Cells["A41"].Value?.ToString(), out int n_value))
                        {
                            numRows1 = n_value;
                        }
                        if (numRows1 <= 0)
                        {
                            Console.WriteLine("Value unvalid.");
                            outputBuilder.AppendLine($"Value unvalid.");
                            File.WriteAllText(outputMixer, outputBuilder.ToString());
<<<<<<< HEAD
                            MessageBox.Show("Value at A41 is unvalid or null. Model Mixer can not be converted", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
=======
>>>>>>> 2742ae2f9991af8eb6b861e2a4bb4cf034f1f554
                            return;
                        }
                        int startRow1 = 42;
                        int counter1 = 1;

                        // 3. Lặp qua từng ô trong ma trận đã xác định
                        for (int row1 = startRow1; row1 < startRow1 + numRows1; row1++)
                        {
                            // Lấy giá trị của ô hiện tại
                            var cellValue1 = worksheet.Cells[row1, 1].Value;
                            // 4. Kiểm tra xem ô có dữ liệu hay không (không rỗng, không null)
                            if (cellValue1 != null && !string.IsNullOrWhiteSpace(cellValue1.ToString()))
                            {
                                // Lấy giá trị input từ cột A cùng hàng
                                var inputValue = worksheet.Cells[row1, 4].Value?.ToString();
                                // Xử lý định dạng số (thay thế 6 số 0 bằng 'e6')
                                // string formattedInput = FormatNumber(inputValue);
                                //kết quả terminal 
                                Console.WriteLine($"SpecMin{counter1}={inputValue}");
                                outputBuilder.AppendLine($"SpecMin{counter1}={inputValue}");
                                counter1++;
                            }
                        }
                        // //   ----------------------------------------- SPECMAX ------------------------------------------------ 
                        int numRows2 = 0;
                        if (int.TryParse(worksheet.Cells["A41"].Value?.ToString(), out int m_value))
                        {
                            numRows2 = m_value;
                        }
                        if (numRows2 <= 0)
                        {
                            Console.WriteLine("Value unvalid.");
                            outputBuilder.AppendLine($"Value unvalid.");
                            File.WriteAllText(outputMixer, outputBuilder.ToString());
<<<<<<< HEAD
                            MessageBox.Show("Value at A41 is unvalid or null. Model Mixer can not be converted", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
=======
>>>>>>> 2742ae2f9991af8eb6b861e2a4bb4cf034f1f554
                            return;
                        }
                        int startRow2 = 42;
                        int counter2 = 1;

                        // 3. Lặp qua từng ô trong ma trận đã xác định
                        for (int row2 = startRow2; row2 < startRow2 + numRows2; row2++)
                        {
                            // Lấy giá trị của ô hiện tại
                            var cellValue2 = worksheet.Cells[row2, 1].Value;
                            // 4. Kiểm tra xem ô có dữ liệu hay không (không rỗng, không null)
                            if (cellValue2 != null && !string.IsNullOrWhiteSpace(cellValue2.ToString()))
                            {
                                // Lấy giá trị input từ cột A cùng hàng
                                // var inputValue = worksheet.Cells[row1, 4].Value?.ToString();
                                // Xử lý định dạng số (thay thế 6 số 0 bằng 'e6')
                                // string formattedInput = FormatNumber(inputValue);
                                //kết quả terminal 
                                Console.WriteLine($"SpecMax{counter2}=99");
                                outputBuilder.AppendLine($"SpecMax{counter2}=99");
                                counter2++;
                            }
                        }
                    }    // [ISO_SPEC]
                    File.WriteAllText(output, outputBuilder.ToString());
                    MessageBox.Show($"Processing Mixer completed successfully. File saved to {outputMixer}", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected Error: {ex.Message}");
                MessageBox.Show($"Error processing Mixer file INI: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
        }



        //-------------------------------------------------------------XỬ LÝ DOUBLER ---------------------------------------------------------------
        private void ProcessModelDoubler()
        {
            UpdateStatus($"Processing Doubler model file...");
            System.Threading.Thread.Sleep(1500); // Giả lập thời gian xử lý
            try
            {
                // create file ini
                string timestamp = DateTime.Now.ToString("HH-mm-ss__dd-MM-yyyy");
<<<<<<< HEAD
                string outputFileName = $"config_Doubler_{timestamp}.ini";
=======
                string outputFileName = $"config_{timestamp}.ini";
>>>>>>> 2742ae2f9991af8eb6b861e2a4bb4cf034f1f554
                string output = Path.Combine(outputDoubler, outputFileName);
                // StringBuilder for ini results
                var outputBuilder = new StringBuilder();
                var fileInfo = new FileInfo(selectedFilePath);
                using (var package = new ExcelPackage(fileInfo))
                {
                    // check worksheet 
                    var worksheet = package.Workbook.Worksheets[0];
                    if (worksheet == null)
                    {
                        Console.WriteLine("Can not find any worksheet in file Excel.");
<<<<<<< HEAD
                        MessageBox.Show("Can not find any worksheet in file Excel.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
=======
>>>>>>> 2742ae2f9991af8eb6b861e2a4bb4cf034f1f554
                        return;
                    }

                    {
                        Console.WriteLine($"[INFO]\n");
                        outputBuilder.AppendLine($"[INFO]\n");
                        //--------------------------------------NumCLStep= -------------------------------------------
                        Console.WriteLine($"NumCLStep=0 ");
                        outputBuilder.AppendLine($"NumCLStep=0 ");

                        //--------------------------------------NumToiStep---------------------------------------------
                        Console.WriteLine($"NumToiStep=0 ");
                        outputBuilder.AppendLine($"NumToiStep=0 ");
                        //--------------------------------------NumIsoStep---------------------------------------------
                        Console.WriteLine($"NumIsoStep=0 ");
                        outputBuilder.AppendLine($"NumIsoStep=0 ");
                        // --------------------------------------NumDblStep---------------------------------------------
                        int numRows2 = 0;
                        if (int.TryParse(worksheet.Cells["A61"].Value?.ToString(), out int f_value))
                        {
                            numRows2 = f_value;
                        }
                        if (numRows2 <= 0)
                        {
                            Console.WriteLine("Value unvalid.");
<<<<<<< HEAD
                            outputBuilder.AppendLine($"Value at A61 is unvalid or null.");
                            File.WriteAllText(outputDoubler, outputBuilder.ToString());
                            MessageBox.Show("Value at A61 is unvalid or null. Model Doubler can not be converted.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
=======
                            outputBuilder.AppendLine($"Value unvalid.");
                            File.WriteAllText(outputDoubler, outputBuilder.ToString());
>>>>>>> 2742ae2f9991af8eb6b861e2a4bb4cf034f1f554
                            return;
                        }
                        int startRow2 = 62;
                        int step2 = 0;
                        // 3. Lặp qua từng ô trong ma trận đã xác định
                        for (int row2 = startRow2; row2 < startRow2 + numRows2; row2++)
                        {
                            // Lấy giá trị của ô hiện tại
                            var cellValue2 = worksheet.Cells[row2, 1].Value;
                            // 4. Kiểm tra xem ô có dữ liệu hay không (không rỗng, không null)
                            if (cellValue2 != null && !string.IsNullOrWhiteSpace(cellValue2.ToString()))
                            {
                                step2++;
                            }
                        }
                        Console.WriteLine($"NumDblStep={step2}");
                        outputBuilder.AppendLine($"NumDblStep={step2}");
                        // ---------------------------------Comment------------------------------------------------
                        var cellValue4 = worksheet.Cells["E59"].Value;
                        if (cellValue4 != null && !string.IsNullOrWhiteSpace(cellValue4.ToString()))
                        {
                            Console.WriteLine($"Comment= {cellValue4}");
                            outputBuilder.AppendLine($"Comment= {cellValue4}");
                        }
                        else
                        {
                            Console.WriteLine($"Comment= ");
                            outputBuilder.AppendLine($"Comment= ");
                        }
                        // ---------------------------------CAL_date------------------------------------------------
                        string date = DateTime.Now.ToString("dd-MM-yyyy");

                        Console.WriteLine($"CAL_DATE= {date} ");
                        outputBuilder.AppendLine($"CAL_DATE= {date}");


                    }      // INFO 
                    {
                        Console.WriteLine($"\n");
                        Console.WriteLine($"[Doubler]\n");
                        outputBuilder.AppendLine($"\n");
                        outputBuilder.AppendLine($"[Doubler]\n");
                        //-------------------------------FREQUENCY ----------------------------------------
                        int numRows1 = 0;
                        if (int.TryParse(worksheet.Cells["A61"].Value?.ToString(), out int n_value))
                        {
                            numRows1 = n_value;
                        }
                        if (numRows1 <= 0)
                        {
                            Console.WriteLine("Value unvalid.");
<<<<<<< HEAD
                            outputBuilder.AppendLine($"Value at A61 is unvalid or null. Model Doubler can not be converted.");
                            File.WriteAllText(outputDoubler, outputBuilder.ToString());
                            MessageBox.Show("Value at A61 is unvalid or null. Model Doubler can not be converted.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
=======
                            outputBuilder.AppendLine($"Value unvalid.");
                            File.WriteAllText(outputDoubler, outputBuilder.ToString());
>>>>>>> 2742ae2f9991af8eb6b861e2a4bb4cf034f1f554
                            return;
                        }
                        int startRow1 = 62;
                        int counter1 = 1;
                        // Hàm xử lý định dạng số (thay thế 6 số 0 cuối bằng 'e6')
                        string FormatNumber(string number)
                        {
                            if (string.IsNullOrEmpty(number))
                                return "N/A";

                            // Kiểm tra nếu chuỗi kết thúc bằng 6 số 0
                            if (number.EndsWith("000000"))
                            {
                                return number.Substring(0, number.Length - 6) + "e6";
                            }
                            return number;
                        }
                        // 3. Lặp qua từng ô trong ma trận đã xác định
                        for (int row1 = startRow1; row1 < startRow1 + numRows1; row1++)
                        {
                            // Lấy giá trị của ô hiện tại
                            var cellValue1 = worksheet.Cells[row1, 1].Value;
                            // 4. Kiểm tra xem ô có dữ liệu hay không (không rỗng, không null)
                            if (cellValue1 != null && !string.IsNullOrWhiteSpace(cellValue1.ToString()))
                            {
                                // Lấy giá trị input từ cột B  cùng hàng
                                var inputValue1 = worksheet.Cells[row1, 1].Value?.ToString();
                                // Xử lý định dạng số (thay thế 6 số 0 bằng 'e6')
                                string formattedInput = FormatNumber(inputValue1);
                                //kết quả terminal 
                                Console.WriteLine($"Freq{counter1}={formattedInput}");
                                outputBuilder.AppendLine($"Freq{counter1}={formattedInput}");
                                counter1++;
                            }
                        }
                        // ---------------------------------POWER---------------------------------------------
                        int numRows2 = 0;
                        if (int.TryParse(worksheet.Cells["A61"].Value?.ToString(), out int m_value))
                        {
                            numRows2 = m_value;
                        }
                        if (numRows2 <= 0)
                        {
                            Console.WriteLine("Value unvalid.");
                            outputBuilder.AppendLine($"Value unvalid.");
                            File.WriteAllText(outputDoubler, outputBuilder.ToString());
                            return;
                        }
                        int startRow2 = 62;
                        int counter2 = 1;


                        // 3. Lặp qua từng ô trong ma trận đã xác định
                        for (int row2 = startRow2; row2 < startRow2 + numRows2; row2++)
                        {
                            // Lấy giá trị của ô hiện tại
                            var cellValue2 = worksheet.Cells[row2, 2].Value;
                            // 4. Kiểm tra xem ô có dữ liệu hay không (không rỗng, không null)
                            if (cellValue2 != null && !string.IsNullOrWhiteSpace(cellValue2.ToString()))
                            {
                                // Lấy giá trị input từ cột B  cùng hàng
                                var inputValue2 = worksheet.Cells[row2, 2].Value?.ToString();
                                // Xử lý định dạng số (thay thế 6 số 0 bằng 'e6')
                                //string formattedInput = FormatNumber(inputValue1);
                                //kết quả terminal 
                                Console.WriteLine($"Power{counter2}={inputValue2}");
                                outputBuilder.AppendLine($"Power{counter2}={inputValue2}");
                                counter2++;
                            }
                        }
                    }    // Doubler
                    {
                        Console.WriteLine($"\n");
                        Console.WriteLine($"[Doubler_RF]\n");
                        outputBuilder.AppendLine($"\n");
                        outputBuilder.AppendLine($"[Doubler_RF]\n");
                        //-------------------------------Doubler_RF ----------------------------------------
                        int numRows1 = 0;
                        if (int.TryParse(worksheet.Cells["A61"].Value?.ToString(), out int n_value))
                        {
                            numRows1 = n_value;
                        }
                        if (numRows1 <= 0)
                        {
                            Console.WriteLine("Value unvalid.");
<<<<<<< HEAD
                            outputBuilder.AppendLine($"Value at A61 is unvalid or null. Model Doubler can not be converted.");
                            File.WriteAllText(outputDoubler, outputBuilder.ToString());
                            MessageBox.Show("Value at A61 is unvalid or null. Model Doubler can not be converted.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
=======
                            outputBuilder.AppendLine($"Value unvalid.");
                            File.WriteAllText(outputDoubler, outputBuilder.ToString());
>>>>>>> 2742ae2f9991af8eb6b861e2a4bb4cf034f1f554
                            return;
                        }
                        int startRow1 = 62;
                        int counter1 = 1;
                        for (int row1 = startRow1; row1 < startRow1 + numRows1; row1++)
                        {
                            // Lấy giá trị của ô hiện tại
                            var cellValue1 = worksheet.Cells[row1, 1].Value;
                            // 4. Kiểm tra xem ô có dữ liệu hay không (không rỗng, không null)
                            if (cellValue1 != null && !string.IsNullOrWhiteSpace(cellValue1.ToString()))
                            {
                                // Lấy giá trị input từ cột B  cùng hàng
                                //var inputValue1 = worksheet.Cells[row1, 1].Value?.ToString();
                                // Xử lý định dạng số (thay thế 6 số 0 bằng 'e6')
                                //string formattedInput = FormatNumber(inputValue1);
                                //kết quả terminal 
                                Console.WriteLine($"SetPower{counter1}=0 ");
                                outputBuilder.AppendLine($"SetPower{counter1}=0 ");
                                counter1++;
                            }
                        }
                    }    // Doubler_RF
                    {
                        Console.WriteLine($"\n");
                        Console.WriteLine($"[Doubler_LO]\n");
                        outputBuilder.AppendLine($"\n");
                        outputBuilder.AppendLine($"[Doubler_LO]\n");
                        //-------------------------------Doubler_RF ----------------------------------------
                        int numRows1 = 0;
                        if (int.TryParse(worksheet.Cells["A61"].Value?.ToString(), out int n_value))
                        {
                            numRows1 = n_value;
                        }
                        if (numRows1 <= 0)
                        {
                            Console.WriteLine("Value unvalid.");
<<<<<<< HEAD
                            outputBuilder.AppendLine($"Value at A61 is unvalid or null. Model Doubler can not be converted.");
                            File.WriteAllText(outputDoubler, outputBuilder.ToString());
                            MessageBox.Show("Value at A61 is unvalid or null. Model Doubler can not be converted.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
=======
                            outputBuilder.AppendLine($"Value unvalid.");
                            File.WriteAllText(outputDoubler, outputBuilder.ToString());
>>>>>>> 2742ae2f9991af8eb6b861e2a4bb4cf034f1f554
                            return;
                        }
                        int startRow1 = 62;
                        int counter1 = 1;
                        for (int row1 = startRow1; row1 < startRow1 + numRows1; row1++)
                        {
                            // Lấy giá trị của ô hiện tại
                            var cellValue1 = worksheet.Cells[row1, 1].Value;
                            // 4. Kiểm tra xem ô có dữ liệu hay không (không rỗng, không null)
                            if (cellValue1 != null && !string.IsNullOrWhiteSpace(cellValue1.ToString()))
                            {
                                // Lấy giá trị input từ cột B  cùng hàng
                                //var inputValue1 = worksheet.Cells[row1, 1].Value?.ToString();
                                // Xử lý định dạng số (thay thế 6 số 0 bằng 'e6')
                                //string formattedInput = FormatNumber(inputValue1);
                                //kết quả terminal 
                                Console.WriteLine($"SetPower{counter1}=0 ");
                                outputBuilder.AppendLine($"SetPower{counter1}=0 ");
                                counter1++;
                            }
                        }
                    }      // Doubler_LO
                    {
                        Console.WriteLine($"\n");
                        Console.WriteLine($"[Doubler_IF1]\n");
                        outputBuilder.AppendLine($"\n");
                        outputBuilder.AppendLine($"[Doubler_IF1]\n");
                        //-------------------------------SPAN----------------------------------------
                        int numRows1 = 0;
                        if (int.TryParse(worksheet.Cells["A61"].Value?.ToString(), out int n_value))
                        {
                            numRows1 = n_value;
                        }
                        if (numRows1 <= 0)
                        {
                            Console.WriteLine("Value unvalid.");
<<<<<<< HEAD
                            outputBuilder.AppendLine($"Value at A61 is unvalid or null. Model Doubler can not be converted.");
                            File.WriteAllText(outputDoubler, outputBuilder.ToString());
                            MessageBox.Show("Value at A61 is unvalid or null. Model Doubler can not be converted.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
=======
                            outputBuilder.AppendLine($"Value unvalid.");
                            File.WriteAllText(outputDoubler, outputBuilder.ToString());
>>>>>>> 2742ae2f9991af8eb6b861e2a4bb4cf034f1f554
                            return;
                        }
                        int startRow1 = 62;
                        int counter1 = 1;
                        // Hàm xử lý định dạng số (thay thế 4 số 0 cuối bằng 'e4')
                        string FormatNumber1(string number)
                        {
                            if (string.IsNullOrEmpty(number))
                                return "N/A";

                            // Kiểm tra nếu chuỗi kết thúc bằng 6 số 0
                            if (number.EndsWith("0000"))
                            {
                                return number.Substring(0, number.Length - 4) + "e4";
                            }
                            return number;
                        }

                        for (int row1 = startRow1; row1 < startRow1 + numRows1; row1++)
                        {
                            // Lấy giá trị của ô hiện tại
                            var cellValue1 = worksheet.Cells[row1, 1].Value;
                            // 4. Kiểm tra xem ô có dữ liệu hay không (không rỗng, không null)
                            if (cellValue1 != null && !string.IsNullOrWhiteSpace(cellValue1.ToString()))
                            {
                                var inputValue1 = worksheet.Cells[row1, 6].Value?.ToString();

                                string formattedInput = FormatNumber1(inputValue1);

                                Console.WriteLine($"Span{counter1}={formattedInput}");
                                outputBuilder.AppendLine($"Span{counter1}={formattedInput}");
                                counter1++;
                            }
                        }
                        //---------------------------------OFFSET-----------------------------------------
                        int numRows2 = 0;
                        if (int.TryParse(worksheet.Cells["A61"].Value?.ToString(), out int m_value))
                        {
                            numRows2 = m_value;
                        }
                        if (numRows2 <= 0)
                        {
                            Console.WriteLine("Value unvalid.");
                            outputBuilder.AppendLine($"Value unvalid.");
                            File.WriteAllText(outputDoubler, outputBuilder.ToString());
                            return;
                        }
                        int startRow2 = 62;
                        int counter2 = 1;
                        for (int row2 = startRow2; row2 < startRow2 + numRows2; row2++)
                        {
                            var cellValue2 = worksheet.Cells[row2, 1].Value;
                            if (cellValue2 != null && !string.IsNullOrWhiteSpace(cellValue2.ToString()))
                            {
                                // var inputValue2 = worksheet.Cells[row2, 7].Value?.ToString();
                                // Xử lý định dạng số (thay thế 4 số 0 bằng 'e4')
                                //string formattedInput = FormatNumber1(inputValue2);
                                Console.WriteLine($"Offset{counter2}=0 ");
                                outputBuilder.AppendLine($"Offset{counter2}=0 ");
                                counter2++;
                            }
                        }
                        //---------------------------------SetPower---------------------------------------
                        int numRows3 = 0;
                        if (int.TryParse(worksheet.Cells["A61"].Value?.ToString(), out int p_value))
                        {
                            numRows3 = p_value;
                        }
                        if (numRows3 <= 0)
                        {
                            Console.WriteLine("Value unvalid.");
                            outputBuilder.AppendLine($"Value unvalid.");
                            File.WriteAllText(outputDoubler, outputBuilder.ToString());
                            return;
                        }
                        int startRow3 = 62;
                        int counter3 = 1;
                        for (int row3 = startRow3; row3 < startRow3 + numRows3; row3++)
                        {
                            // Lấy giá trị của ô hiện tại
                            var cellValue3 = worksheet.Cells[row3, 1].Value;
                            // 4. Kiểm tra xem ô có dữ liệu hay không (không rỗng, không null)
                            if (cellValue3 != null && !string.IsNullOrWhiteSpace(cellValue3.ToString()))
                            {
                                // Lấy giá trị input từ cột B  cùng hàng
                                //var inputValue3 = worksheet.Cells[row3, 8].Value?.ToString();
                                // Xử lý định dạng số (thay thế 4 số 0 bằng 'e4')
                                //string formattedInput = FormatNumber1(inputValue3);
                                Console.WriteLine($"SetPower{counter3}=0 ");
                                outputBuilder.AppendLine($"SetPower{counter3}=0 ");
                                counter3++;
                            }
                        }


                    }      // Doubler_IF1
                    {
                        Console.WriteLine($"\n");
                        Console.WriteLine($"[Doubler_IF2]\n");
                        outputBuilder.AppendLine($"\n");
                        outputBuilder.AppendLine($"[Doubler_IF2]\n");
                        //-------------------------------SPAN----------------------------------------
                        int numRows1 = 0;
                        if (int.TryParse(worksheet.Cells["A61"].Value?.ToString(), out int n_value))
                        {
                            numRows1 = n_value;
                        }
                        if (numRows1 <= 0)
                        {
                            Console.WriteLine("Value unvalid.");
<<<<<<< HEAD
                            outputBuilder.AppendLine($"Value at A61 is unvalid or null. Model Doubler can not be converted.");
                            File.WriteAllText(outputDoubler, outputBuilder.ToString());
                            MessageBox.Show("Value at A61 is unvalid or null. Model Doubler can not be converted.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
=======
                            outputBuilder.AppendLine($"Value unvalid.");
                            File.WriteAllText(outputDoubler, outputBuilder.ToString());
>>>>>>> 2742ae2f9991af8eb6b861e2a4bb4cf034f1f554
                            return;
                        }
                        int startRow1 = 62;
                        int counter1 = 1;
                        // Hàm xử lý định dạng số (thay thế 4 số 0 cuối bằng 'e4')
                        string FormatNumber1(string number)
                        {
                            if (string.IsNullOrEmpty(number))
                                return "N/A";

                            // Kiểm tra nếu chuỗi kết thúc bằng 6 số 0
                            if (number.EndsWith("0000"))
                            {
                                return number.Substring(0, number.Length - 4) + "e4";
                            }
                            return number;
                        }

                        for (int row1 = startRow1; row1 < startRow1 + numRows1; row1++)
                        {
                            // Lấy giá trị của ô hiện tại
                            var cellValue1 = worksheet.Cells[row1, 1].Value;
                            // 4. Kiểm tra xem ô có dữ liệu hay không (không rỗng, không null)
                            if (cellValue1 != null && !string.IsNullOrWhiteSpace(cellValue1.ToString()))
                            {
                                var inputValue1 = worksheet.Cells[row1, 7].Value?.ToString();

                                string formattedInput = FormatNumber1(inputValue1);

                                Console.WriteLine($"Span{counter1}={formattedInput}");
                                outputBuilder.AppendLine($"Span{counter1}={formattedInput}");
                                counter1++;
                            }
                        }
                        //---------------------------------OFFSET-----------------------------------------
                        int numRows2 = 0;
                        if (int.TryParse(worksheet.Cells["A61"].Value?.ToString(), out int m_value))
                        {
                            numRows2 = m_value;
                        }
                        if (numRows2 <= 0)
                        {
                            Console.WriteLine("Value unvalid.");
                            outputBuilder.AppendLine($"Value unvalid.");
                            File.WriteAllText(outputDoubler, outputBuilder.ToString());
                            return;
                        }
                        int startRow2 = 62;
                        int counter2 = 1;
                        for (int row2 = startRow2; row2 < startRow2 + numRows2; row2++)
                        {
                            var cellValue2 = worksheet.Cells[row2, 1].Value;
                            if (cellValue2 != null && !string.IsNullOrWhiteSpace(cellValue2.ToString()))
                            {
                                // var inputValue2 = worksheet.Cells[row2, 7].Value?.ToString();
                                // Xử lý định dạng số (thay thế 4 số 0 bằng 'e4')
                                //string formattedInput = FormatNumber1(inputValue2);
                                Console.WriteLine($"Offset{counter2}=0 ");
                                outputBuilder.AppendLine($"Offset{counter2}=0 ");
                                counter2++;
                            }
                        }
                        //---------------------------------SetPower---------------------------------------
                        int numRows3 = 0;
                        if (int.TryParse(worksheet.Cells["A61"].Value?.ToString(), out int p_value))
                        {
                            numRows3 = p_value;
                        }
                        if (numRows3 <= 0)
                        {
                            Console.WriteLine("Value unvalid.");
                            outputBuilder.AppendLine($"Value unvalid.");
                            File.WriteAllText(outputDoubler, outputBuilder.ToString());
                            return;
                        }
                        int startRow3 = 62;
                        int counter3 = 1;
                        for (int row3 = startRow3; row3 < startRow3 + numRows3; row3++)
                        {
                            // Lấy giá trị của ô hiện tại
                            var cellValue3 = worksheet.Cells[row3, 1].Value;
                            // 4. Kiểm tra xem ô có dữ liệu hay không (không rỗng, không null)
                            if (cellValue3 != null && !string.IsNullOrWhiteSpace(cellValue3.ToString()))
                            {
                                // Lấy giá trị input từ cột B  cùng hàng
                                //var inputValue3 = worksheet.Cells[row3, 8].Value?.ToString();
                                // Xử lý định dạng số (thay thế 4 số 0 bằng 'e4')
                                //string formattedInput = FormatNumber1(inputValue3);
                                Console.WriteLine($"SetPower{counter3}=0 ");
                                outputBuilder.AppendLine($"SetPower{counter3}=0 ");
                                counter3++;
                            }
                        }
                    }      // Doubler_IF2
                    {
                        Console.WriteLine($"\n");
                        Console.WriteLine($"[Doubler_IF3]\n");
                        outputBuilder.AppendLine($"\n");
                        outputBuilder.AppendLine($"[Doubler_IF3]\n");
                        //-------------------------------SPAN----------------------------------------
                        int numRows1 = 0;
                        if (int.TryParse(worksheet.Cells["A61"].Value?.ToString(), out int n_value))
                        {
                            numRows1 = n_value;
                        }
                        if (numRows1 <= 0)
                        {
                            Console.WriteLine("Value unvalid.");
<<<<<<< HEAD
                            outputBuilder.AppendLine($"Value at A61 is unvalid or null. Model Doubler can not be converted.");
                            File.WriteAllText(outputDoubler, outputBuilder.ToString());
                            MessageBox.Show("Value at A61 is unvalid or null. Model Doubler can not be converted.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
=======
                            outputBuilder.AppendLine($"Value unvalid.");
                            File.WriteAllText(outputDoubler, outputBuilder.ToString());
>>>>>>> 2742ae2f9991af8eb6b861e2a4bb4cf034f1f554
                            return;
                        }
                        int startRow1 = 62;
                        int counter1 = 1;
                        // Hàm xử lý định dạng số (thay thế 4 số 0 cuối bằng 'e4')
                        string FormatNumber1(string number)
                        {
                            if (string.IsNullOrEmpty(number))
                                return "N/A";

                            // Kiểm tra nếu chuỗi kết thúc bằng 6 số 0
                            if (number.EndsWith("0000"))
                            {
                                return number.Substring(0, number.Length - 4) + "e4";
                            }
                            return number;
                        }

                        for (int row1 = startRow1; row1 < startRow1 + numRows1; row1++)
                        {
                            // Lấy giá trị của ô hiện tại
                            var cellValue1 = worksheet.Cells[row1, 1].Value;
                            // 4. Kiểm tra xem ô có dữ liệu hay không (không rỗng, không null)
                            if (cellValue1 != null && !string.IsNullOrWhiteSpace(cellValue1.ToString()))
                            {
                                var inputValue1 = worksheet.Cells[row1, 8].Value?.ToString();

                                string formattedInput = FormatNumber1(inputValue1);

                                Console.WriteLine($"Span{counter1}={formattedInput}");
                                outputBuilder.AppendLine($"Span{counter1}={formattedInput}");
                                counter1++;
                            }
                        }
                        //---------------------------------OFFSET-----------------------------------------
                        int numRows2 = 0;
                        if (int.TryParse(worksheet.Cells["A61"].Value?.ToString(), out int m_value))
                        {
                            numRows2 = m_value;
                        }
                        if (numRows2 <= 0)
                        {
                            Console.WriteLine("Value unvalid.");
                            outputBuilder.AppendLine($"Value unvalid.");
                            File.WriteAllText(outputDoubler, outputBuilder.ToString());
                            return;
                        }
                        int startRow2 = 62;
                        int counter2 = 1;
                        for (int row2 = startRow2; row2 < startRow2 + numRows2; row2++)
                        {
                            var cellValue2 = worksheet.Cells[row2, 1].Value;
                            if (cellValue2 != null && !string.IsNullOrWhiteSpace(cellValue2.ToString()))
                            {
                                // var inputValue2 = worksheet.Cells[row2, 7].Value?.ToString();
                                // Xử lý định dạng số (thay thế 4 số 0 bằng 'e4')
                                //string formattedInput = FormatNumber1(inputValue2);
                                Console.WriteLine($"Offset{counter2}=0 ");
                                outputBuilder.AppendLine($"Offset{counter2}=0 ");
                                counter2++;
                            }
                        }
                        //---------------------------------SetPower---------------------------------------
                        int numRows3 = 0;
                        if (int.TryParse(worksheet.Cells["A61"].Value?.ToString(), out int p_value))
                        {
                            numRows3 = p_value;
                        }
                        if (numRows3 <= 0)
                        {
                            Console.WriteLine("Value unvalid.");
                            outputBuilder.AppendLine($"Value unvalid.");
                            File.WriteAllText(outputDoubler, outputBuilder.ToString());
                            return;
                        }
                        int startRow3 = 62;
                        int counter3 = 1;
                        for (int row3 = startRow3; row3 < startRow3 + numRows3; row3++)
                        {
                            // Lấy giá trị của ô hiện tại
                            var cellValue3 = worksheet.Cells[row3, 1].Value;
                            // 4. Kiểm tra xem ô có dữ liệu hay không (không rỗng, không null)
                            if (cellValue3 != null && !string.IsNullOrWhiteSpace(cellValue3.ToString()))
                            {
                                // Lấy giá trị input từ cột B  cùng hàng
                                //var inputValue3 = worksheet.Cells[row3, 8].Value?.ToString();
                                // Xử lý định dạng số (thay thế 4 số 0 bằng 'e4')
                                //string formattedInput = FormatNumber1(inputValue3);
                                Console.WriteLine($"SetPower{counter3}=0 ");
                                outputBuilder.AppendLine($"SetPower{counter3}=0 ");
                                counter3++;
                            }
                        }
                    }     // Doubler_IF3
                    {
                        Console.WriteLine($"\n");
                        Console.WriteLine($"[Doubler_Spec]\n");
                        outputBuilder.AppendLine($"\n");
                        outputBuilder.AppendLine($"[Doubler_Spec]\n");
                        //-------------------------------F1_MAX----------------------------------------
                        int numRows1 = 0;
                        if (int.TryParse(worksheet.Cells["A61"].Value?.ToString(), out int n_value))
                        {
                            numRows1 = n_value;
                        }
                        if (numRows1 <= 0)
                        {
                            Console.WriteLine("Value unvalid.");
<<<<<<< HEAD
                            outputBuilder.AppendLine($"Value at A61 is unvalid or null. Model Doubler can not be converted.");
                            File.WriteAllText(outputDoubler, outputBuilder.ToString());
                            MessageBox.Show("Value at A61 is unvalid or null. Model Doubler can not be converted.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
=======
                            outputBuilder.AppendLine($"Value unvalid.");
                            File.WriteAllText(outputDoubler, outputBuilder.ToString());
>>>>>>> 2742ae2f9991af8eb6b861e2a4bb4cf034f1f554
                            return;
                        }
                        int startRow1 = 62;
                        int counter1 = 1;
                        double ParseScientificNumber(string sciNumber)
                        {
                            if (string.IsNullOrEmpty(sciNumber))
                                return 0;

                            // Xử lý chuỗi dạng "1410e6"
                            if (sciNumber.Contains("e") || sciNumber.Contains("E"))
                            {
                                // Tách phần cơ số và số mũ
                                char[] separators = new char[] { 'e', 'E' };
                                string[] parts = sciNumber.Split(separators, 2);

                                if (parts.Length == 2)
                                {
                                    if (double.TryParse(parts[0], out double baseValue) &&
                                        double.TryParse(parts[1], out double exponent))
                                    {
                                        return baseValue * Math.Pow(10, exponent);
                                    }
                                }
                            }

                            // Xử lý chuỗi số thông thường
                            if (double.TryParse(sciNumber, out double result))
                                return result;

                            // Trả về 0 nếu không thể chuyển đổi
                            return 0;
                        }

                        for (int row1 = startRow1; row1 < startRow1 + numRows1; row1++)
                        {
                            // Lấy giá trị của ô hiện tại
                            var cellValue1 = worksheet.Cells[row1, 1].Value;
                            // 4. Kiểm tra xem ô có dữ liệu hay không (không rỗng, không null)
                            if (cellValue1 != null && !string.IsNullOrWhiteSpace(cellValue1.ToString()))
                            {
                                var inputValue1 = worksheet.Cells[row1, 3].Value?.ToString();
                                double specmin = ParseScientificNumber(inputValue1);
                                double result = -specmin;

                                //string formattedInput = FormatNumber1(inputValue1);

                                Console.WriteLine($"F1_Max{counter1}= {result}");
                                outputBuilder.AppendLine($"F1_Max{counter1}= {result}");
                                counter1++;
                            }
                        }
                        //---------------------------------F1_MIN---------------------------------------
                        int numRows2 = 0;
                        if (int.TryParse(worksheet.Cells["A61"].Value?.ToString(), out int m_value))
                        {
                            numRows2 = m_value;
                        }
                        if (numRows2 <= 0)
                        {
                            Console.WriteLine("Value unvalid.");
                            outputBuilder.AppendLine($"Value unvalid.");
                            File.WriteAllText(outputDoubler, outputBuilder.ToString());
                            return;
                        }
                        int startRow2 = 62;
                        int counter2 = 1;
                        for (int row2 = startRow2; row2 < startRow2 + numRows2; row2++)
                        {
                            // Lấy giá trị của ô hiện tại
                            var cellValue2 = worksheet.Cells[row2, 1].Value;
                            // 4. Kiểm tra xem ô có dữ liệu hay không (không rỗng, không null)
                            if (cellValue2 != null && !string.IsNullOrWhiteSpace(cellValue2.ToString()))
                            {
                                //var inputValue2  = worksheet.Cells[row2 , 3].Value?.ToString();
                                Console.WriteLine($"F1_Min{counter2}=-99");
                                outputBuilder.AppendLine($"F1_Min{counter2}=-99");
                                counter2++;
                            }
                        }
                        // ---------------------------------F2_MAX---------------------------------------
                        int numRows3 = 0;
                        if (int.TryParse(worksheet.Cells["A61"].Value?.ToString(), out int p_value))
                        {
                            numRows3 = p_value;
                        }
                        if (numRows3 <= 0)
                        {
                            Console.WriteLine("Value unvalid.");
                            outputBuilder.AppendLine($"Value unvalid.");
                            File.WriteAllText(outputDoubler, outputBuilder.ToString());
                            return;
                        }
                        int startRow3 = 62;
                        int counter3 = 1;
                        for (int row3 = startRow3; row3 < startRow3 + numRows3; row3++)
                        {
                            // Lấy giá trị của ô hiện tại
                            var cellValue3 = worksheet.Cells[row3, 1].Value;
                            // 4. Kiểm tra xem ô có dữ liệu hay không (không rỗng, không null)
                            if (cellValue3 != null && !string.IsNullOrWhiteSpace(cellValue3.ToString()))
                            {
                                //var inputValue2  = worksheet.Cells[row2 , 3].Value?.ToString();
                                Console.WriteLine($"F2_Max{counter3}=-3");
                                outputBuilder.AppendLine($"F2_Max{counter3}=-3");
                                counter3++;
                            }
                        }
                        // ---------------------------------F2_MIN---------------------------------------
                        int numRows4 = 0;
                        if (int.TryParse(worksheet.Cells["A61"].Value?.ToString(), out int h_value))
                        {
                            numRows4 = h_value;
                        }
                        if (numRows4 <= 0)
                        {
                            Console.WriteLine("Value unvalid.");
                            outputBuilder.AppendLine($"Value unvalid.");
                            File.WriteAllText(outputDoubler, outputBuilder.ToString());
                            return;
                        }
                        int startRow4 = 62;
                        int counter4 = 1;
                        for (int row4 = startRow4; row4 < startRow4 + numRows4; row4++)
                        {
                            // Lấy giá trị của ô hiện tại
                            var cellValue4 = worksheet.Cells[row4, 1].Value;
                            // 4. Kiểm tra xem ô có dữ liệu hay không (không rỗng, không null)
                            if (cellValue4 != null && !string.IsNullOrWhiteSpace(cellValue4.ToString()))
                            {
                                var inputValue2 = worksheet.Cells[row4, 4].Value?.ToString();
                                Console.WriteLine($"F2_Min{counter4}=-{inputValue2}");
                                outputBuilder.AppendLine($"F2_Min{counter4}=-{inputValue2}");
                                counter4++;
                            }
                        }
                        // ---------------------------------F3_MIN---------------------------------------
                        int numRows5 = 0;
                        if (int.TryParse(worksheet.Cells["A61"].Value?.ToString(), out int f_value))
                        {
                            numRows5 = f_value;
                        }
                        if (numRows5 <= 0)
                        {
                            Console.WriteLine("Value unvalid.");
                            outputBuilder.AppendLine($"Value unvalid.");
                            File.WriteAllText(outputDoubler, outputBuilder.ToString());
                            return;
                        }
                        int startRow5 = 62;
                        int counter5 = 1;
                        for (int row5 = startRow5; row5 < startRow5 + numRows5; row5++)
                        {
                            // Lấy giá trị của ô hiện tại
                            var cellValue5 = worksheet.Cells[row5, 1].Value;
                            // 4. Kiểm tra xem ô có dữ liệu hay không (không rỗng, không null)
                            if (cellValue5 != null && !string.IsNullOrWhiteSpace(cellValue5.ToString()))
                            {
                                //var inputValue2 = worksheet.Cells[row5, 4].Value?.ToString();
                                Console.WriteLine($"F3_Min{counter5}=-99");
                                outputBuilder.AppendLine($"F3_Min{counter5}=-99");
                                counter5++;
                            }
                        }
                        // ---------------------------------F3_MAX---------------------------------------
                        int numRows6 = 0;
                        if (int.TryParse(worksheet.Cells["A61"].Value?.ToString(), out int j_value))
                        {
                            numRows6 = j_value;
                        }
                        if (numRows6 <= 0)
                        {
                            Console.WriteLine("Value unvalid.");
                            outputBuilder.AppendLine($"Value unvalid.");
                            File.WriteAllText(outputDoubler, outputBuilder.ToString());
                            return;
                        }
                        int startRow6 = 62;
                        int counter6 = 1;
                        for (int row6 = startRow6; row6 < startRow6 + numRows6; row6++)
                        {
                            // Lấy giá trị của ô hiện tại
                            var cellValue6 = worksheet.Cells[row6, 1].Value;
                            // 4. Kiểm tra xem ô có dữ liệu hay không (không rỗng, không null)
                            if (cellValue6 != null && !string.IsNullOrWhiteSpace(cellValue6.ToString()))
                            {
                                var inputValue2 = worksheet.Cells[row6, 5].Value?.ToString();
                                Console.WriteLine($"F3_Max{counter6}=-{inputValue2}");
                                outputBuilder.AppendLine($"F3_Max{counter6}=-{inputValue2}");
                                counter6++;
                            }
                        }
                    }     // Double_Spec
                    File.WriteAllText(output, outputBuilder.ToString());
                    MessageBox.Show($"Processing Doubler completed successfully. File saved to {outputDoubler}", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error processing Doubler file INI: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
        }


<<<<<<< HEAD
=======


>>>>>>> 2742ae2f9991af8eb6b861e2a4bb4cf034f1f554
    }
}
