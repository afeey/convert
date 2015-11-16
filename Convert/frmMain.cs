using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Drawing.Imaging;

using Microsoft.Office.Interop.Word;
using Microsoft.Office.Interop.Excel;

namespace Convert
{
    public partial class frmMain : Form
    {
        public frmMain()
        {
            InitializeComponent();
        }

        private void btnSelect_Click(object sender, EventArgs e)
        {
            DialogResult result = selectWordDialog.ShowDialog();
            if (result == DialogResult.OK)
            {
                string[] fileNames = selectWordDialog.FileNames;
                foreach (string fileName in fileNames)
                {
                    lbFiles.Items.Add(fileName);
                }
            }

        }

        private void btnClear_Click(object sender, EventArgs e)
        {
            lbFiles.Items.Clear();
        }

        private void btnConvert_Click(object sender, EventArgs e)
        {
            btnSelect.Enabled = false;
            btnClear.Enabled = false;
            btnConvert.Enabled = false;

            lblMsg.Text = "正在转换，请稍后...";
            for (int i = 0; i < lbFiles.Items.Count; i++)
            {
                string strFailTables = "";
                string strNoPhoto = "";
                bool b = DataConvert(lbFiles.Items[i].ToString(), out strFailTables, out strNoPhoto);

                if (!b)
                {
                    lbFiles.Items[i] = lbFiles.Items[i].ToString() + "    失败";
                    continue;
                }

                if (strFailTables.Length > 0)
                {
                    lbFiles.Items[i] = lbFiles.Items[i].ToString() + "    完成  [表" + strFailTables + "数据转换及照片提取失败]";
                }
                else
                {
                    lbFiles.Items[i] = lbFiles.Items[i].ToString() + "    完成";
                }

                if (strNoPhoto.Length > 0)
                {
                    lbFiles.Items[i] = lbFiles.Items[i].ToString() + "[表" + strNoPhoto + "照片提取失败]";
                }
            }
            lblMsg.Text = "转换全部完成";
            btnSelect.Enabled = true;
            btnClear.Enabled = true;
            btnConvert.Enabled = true;
        }


        private bool DataConvert(string fileName, out string failTables, out string noPhotos)
        {
            string fileDir = fileName.Substring(0, fileName.LastIndexOf("\\"));
            string photoDir = fileName.Substring(0, fileName.LastIndexOf(".")) + "_Photos";
            if (Directory.Exists(photoDir))
            {
                Directory.Delete(photoDir, true);
            }
            Directory.CreateDirectory(photoDir);


            bool b = true;

            Microsoft.Office.Interop.Word.ApplicationClass appClass = null;
            Microsoft.Office.Interop.Word.Document doc = null;

            object path = fileName;
            object missing = System.Reflection.Missing.Value;

            int count = 0;
            failTables = "";
            noPhotos = "";
            try
            {
                appClass = new Microsoft.Office.Interop.Word.ApplicationClass();
                appClass.Visible = false;
                doc = appClass.Documents.Open(ref path, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing);

                System.Data.DataTable dt = new System.Data.DataTable();
                dt.Columns.Add("RowNo", typeof(Int32));
                dt.Columns.Add("IDCode", typeof(String));           //身份证编号
                dt.Columns.Add("Name", typeof(String));             //姓名
                dt.Columns.Add("Sex", typeof(String));              //性别
                dt.Columns.Add("Nation", typeof(String));           //民族
                dt.Columns.Add("Education", typeof(String));        //学历
                dt.Columns.Add("Politics", typeof(String));         //政治面貌
                dt.Columns.Add("WorkUnit", typeof(String));         //工作单位
                dt.Columns.Add("Position", typeof(String));         //职务
                dt.Columns.Add("PositionLevel", typeof(String));    //职级
                dt.Columns.Add("JobTitle", typeof(String));         //职称
                dt.Columns.Add("Compilation", typeof(String));      //编制
                dt.Columns.Add("LawClass", typeof(String));         //执法种类
                dt.Columns.Add("LawArea", typeof(String));          //执法区域

                int rowNo = 1;
                foreach (Table table in doc.Tables)
                {
                    count++;
                    try
                    {
                        DataRow row = dt.NewRow();
                        row["RowNo"] = rowNo;
                        row["Name"] = table.Cell(1, 2).Range.Text.Replace("\r\a", "").Trim().Replace(" ", "");
                        row["Sex"] = table.Cell(1, 4).Range.Text.Replace("\r\a", "").Trim();
                        row["Nation"] = table.Cell(2, 2).Range.Text.Replace("\r\a", "").Trim();
                        row["Education"] = table.Cell(2, 4).Range.Text.Replace("\r\a", "").Trim();
                        row["Politics"] = table.Cell(2, 6).Range.Text.Replace("\r\a", "").Trim();
                        row["WorkUnit"] = table.Cell(3, 2).Range.Text.Replace("\r\a", "").Trim();
                        row["Position"] = table.Cell(4, 2).Range.Text.Replace("\r\a", "").Trim();
                        row["PositionLevel"] = table.Cell(4, 4).Range.Text.Replace("\r\a", "").Trim();
                        row["JobTitle"] = table.Cell(4, 6).Range.Text.Replace("\r\a", "").Trim();
                        row["Compilation"] = table.Cell(4, 8).Range.Text.Replace("\r\a", "").Trim();
                        row["LawClass"] = table.Cell(5, 2).Range.Text.Replace("\r\a", "").Trim();
                        row["LawArea"] = table.Cell(5, 4).Range.Text.Replace("\r\a", "").Trim();
                        row["IDCode"] = table.Cell(6, 4).Range.Text.Replace("\r\a", "").Trim();

                        bool photoFlag = false;
                        if (table.Cell(1, 7).Range.InlineShapes.Count != 0)
                        {
                            try
                            {
                                InlineShape shape = table.Cell(1, 7).Range.InlineShapes[1];
                                if (shape.Type == WdInlineShapeType.wdInlineShapePicture || shape.Type == WdInlineShapeType.wdInlineShapeLinkedPicture)
                                {
                                    //利用剪贴板保存数据
                                    shape.Select(); //选定当前图片
                                    appClass.Selection.CopyAsPicture();//copy当前图片
                                    if (Clipboard.ContainsImage())
                                    {
                                        Image img = Clipboard.GetImage();
                                        Bitmap bmp = new Bitmap(img);

                                        string imageName = photoDir + "\\" + row["IDCode"].ToString().Trim() + ".jpg";
                                        int i = 2;
                                        while (File.Exists(imageName))
                                        {
                                            imageName = photoDir + "\\" + row["IDCode"].ToString().Trim() + "(" + i + ").jpg";
                                            i++;
                                        }

                                        EncoderParameters parameters = new EncoderParameters(1);
                                        EncoderParameter parameter = new EncoderParameter(System.Drawing.Imaging.Encoder.Quality,100L);
                                        parameters.Param[0] = parameter;

                                        ImageCodecInfo codecInfo=null;
                                        ImageCodecInfo[] codecInfos = ImageCodecInfo.GetImageEncoders();

                                        codecInfo = codecInfos[1];
                                        bmp.Save(imageName, codecInfos[1], parameters);


                                        //bmp.Save(imageName, System.Drawing.Imaging.ImageFormat.Jpeg);
                                        photoFlag = true;
                                    }
                                }
                            }
                            catch (Exception error)
                            {
                                photoFlag = false;
                            }
                        }
                        if (!photoFlag)
                        {
                            noPhotos += "," + count;
                        }

                        dt.Rows.Add(row);
                        rowNo++;
                    }
                    catch
                    {
                        failTables += "," + count;
                    }
                }

                if (failTables.Length > 0)
                {
                    failTables = failTables.Substring(1);
                }

                if (noPhotos.Length > 0)
                {
                    noPhotos = noPhotos.Substring(1);
                }

                string excelName = fileName.Substring(0, fileName.LastIndexOf(".") + 1) + "xlsx";
                b = DataTableToExcel(dt, excelName);

                if (b)
                {
                    dgvData.DataSource = dt;
                }
                else
                {
                    MessageBox.Show("生成Excel时发生错误", "提示");
                }

            }
            catch (Exception err)
            {
                b = false;
                MessageBox.Show(err.Message, "异常");
            }
            finally
            {
                Clipboard.Clear();
                if (doc != null)
                {
                    doc.Close(ref missing, ref missing, ref missing);
                }
                if (appClass != null)
                {
                    appClass.Quit(ref missing, ref missing, ref missing);
                }

                GC.Collect();
            }
            return b;
        }


        private bool DataTableToExcel(System.Data.DataTable table, string fileName)
        {
            bool b = true;
            Microsoft.Office.Interop.Excel.Application excel = null;
            Workbook workbook = null;
            Worksheet worksheet = null;

            try
            {
                excel = new Microsoft.Office.Interop.Excel.ApplicationClass();
                excel.Visible = false;
                workbook = excel.Workbooks.Add(Microsoft.Office.Interop.Excel.XlWBATemplate.xlWBATWorksheet);
                worksheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Worksheets[1];
                worksheet.Cells.Font.Size = 10;
                worksheet.Cells.NumberFormat = "@";

                int row = 1;
                int column = 1;

                worksheet.Cells[row, column + 1] = "身份证号";
                worksheet.Cells[row, column + 2] = "姓名";
                worksheet.Cells[row, column + 3] = "主体编码";
                worksheet.Cells[row, column + 4] = "单位名称";
                worksheet.Cells[row, column + 5] = "性别";
                worksheet.Cells[row, column + 6] = "政治面貌";
                worksheet.Cells[row, column + 7] = "名族";
                worksheet.Cells[row, column + 8] = "学历";
                worksheet.Cells[row, column + 9] = "职级";
                worksheet.Cells[row, column + 10] = "职务";
                worksheet.Cells[row, column + 11] = "人员类别";
                worksheet.Cells[row, column + 12] = "编制情况";
                worksheet.Cells[row, column + 13] = "执法种类";
                worksheet.Cells[row, column + 14] = "执法区域";
                worksheet.Cells[row, column + 15] = "所属区县";
                worksheet.Cells[row, column + 16] = "照片";
                worksheet.Cells[row, column + 17] = "照片地址";

                worksheet.get_Range(worksheet.Cells[row, column + 1], worksheet.Cells[row, column + 17]).Font.Bold = true;

                for (int i = 1; i <= table.Rows.Count; i++)
                {

                    DataRow dataRow = table.Rows[i - 1];
                    worksheet.Cells[row + i, column + 1] = dataRow["IDCode"].ToString();
                    worksheet.Cells[row + i, column + 2] = dataRow["Name"].ToString();
                    worksheet.Cells[row + i, column + 3] = "";
                    worksheet.Cells[row + i, column + 4] = dataRow["WorkUnit"].ToString();
                    worksheet.Cells[row + i, column + 5] = dataRow["Sex"].ToString();
                    worksheet.Cells[row + i, column + 6] = dataRow["Politics"].ToString();
                    worksheet.Cells[row + i, column + 7] = dataRow["Nation"].ToString();
                    worksheet.Cells[row + i, column + 8] = dataRow["Education"].ToString();
                    worksheet.Cells[row + i, column + 9] = dataRow["PositionLevel"].ToString();
                    worksheet.Cells[row + i, column + 10] = dataRow["Position"].ToString();
                    worksheet.Cells[row + i, column + 11] = "";
                    worksheet.Cells[row + i, column + 12] = dataRow["Compilation"].ToString();
                    worksheet.Cells[row + i, column + 13] = dataRow["LawClass"].ToString();
                    worksheet.Cells[row + i, column + 14] = dataRow["LawArea"].ToString();
                    worksheet.Cells[row + i, column + 15] = dataRow["LawArea"].ToString();
                    worksheet.Cells[row + i, column + 16] = dataRow["IDCode"].ToString() + ".jpg";
                    worksheet.Cells[row + i, column + 17] = "";
                }

                if (File.Exists(fileName))
                {
                    File.Delete(fileName);
                }

                workbook.Saved = true;
                workbook.SaveCopyAs(fileName);
                workbook.Close(false, null, null);
                b = true;

            }
            catch (Exception err)
            {
                MessageBox.Show(err.Message);
                b = false;
            }
            finally
            {
                if (excel != null)
                {
                    excel.Quit();
                    excel = null;
                }
                GC.Collect();
            }

            return b;
        }
    }
}
