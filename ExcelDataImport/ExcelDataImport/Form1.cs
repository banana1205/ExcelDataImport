using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;

namespace ExcelDataImport
{
    public partial class Form1 : Form
    {
        private string table_import = string.Empty;//导入的目标表
        private List<XMLTableColumnInfo> tcL_xml;//需要导入的表模板信息
        private DataTable dt_view = new DataTable();//导入的数据预览
        public Form1()
        {
            InitializeComponent();
        }
        /// <summary>
        /// 窗体加载
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Form1_Load(object sender, EventArgs e)
        {
            LoadXmlTable();
        }

        /// <summary>
        /// 选择Excel文件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btn_openFile_Click(object sender, EventArgs e)
        {
            try
            {
                //获取Excel文件路径和名称  
                OpenFileDialog odXls = new OpenFileDialog();
                //指定相应的打开文档的目录  AppDomain.CurrentDomain.BaseDirectory定位到Debug目录，再根据实际情况进行目录调整  
                string folderPath = AppDomain.CurrentDomain.BaseDirectory + @"databackup\";
                odXls.InitialDirectory = folderPath;
                // 设置文件格式    
                odXls.Filter = "Excel files office2003(*.xls)|*.xls|Excel office2010(*.xlsx)|*.xlsx|All files (*.*)|*.*";
                //openFileDialog1.Filter = "图片文件(*.jpg)|*.jpg|(*.JPEG)|*.jpeg|(*.PNG)|*.png";  
                odXls.FilterIndex = 2;
                odXls.RestoreDirectory = true;
                if (odXls.ShowDialog() == DialogResult.OK)
                {
                    this.txt_filePath.Text = odXls.FileName;
                    this.txt_filePath.ReadOnly = true;
                    string sConnString = string.Format("Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source={0};" + "Extended Properties='Excel 8.0;HDR=NO;IMEX=1';", odXls.FileName);
                    if ((System.IO.Path.GetExtension(txt_filePath.Text.Trim())).ToLower() == ".xls")
                    {
                        sConnString = "Provider=Microsoft.Jet.OLEDB.4.0;" + "data source=" + odXls.FileName + ";Extended Properties=Excel 5.0;Persist Security Info=False";
                        //sConnString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + txtFilePath.Text.Trim() + ";Extended Properties=\"Excel 8.0;HDR=" + strHead + ";IMEX=1\"";  
                    }
                    using (OleDbConnection oleDbConn = new OleDbConnection(sConnString))
                    {
                        oleDbConn.Open();
                        DataTable dt = oleDbConn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" });
                        //判断是否cmb中已有数据，有则清空  
                        if (cmb_sheetName.Items.Count > 0)
                        {
                            cmb_sheetName.DataSource = null;
                            cmb_sheetName.Items.Clear();
                        }
                        //遍历dt的rows得到所有的TABLE_NAME，并Add到cmb中  
                        foreach (DataRow dr in dt.Rows)
                        {
                            cmb_sheetName.Items.Add((String)dr["TABLE_NAME"]);
                        }
                        if (cmb_sheetName.Items.Count > 0)
                        {
                            cmb_sheetName.SelectedIndex = 0;
                        }

                    }
                    //加载Excel文件数据按钮  
                    this.txt_filePath.ReadOnly = true;
                    this.cmb_sheetName.DropDownStyle = ComboBoxStyle.DropDownList;
                    this.cmb_modelList.DropDownStyle = ComboBoxStyle.DropDownList;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        /// <summary>
        /// Excel加载
        /// </summary>
        private void LoadExcel()
        {
            #region 读取相应的表名的Excel文件中数据到当前DataGridview中显示  
            OleDbConnection ole = null;
            OleDbDataAdapter da = null;
            dt_view = new DataTable();
            string strConn = string.Format("Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source={0};" + "Extended Properties='Excel 8.0;HDR=YES;IMEX=1';", txt_filePath.Text.Trim());
            if ((System.IO.Path.GetExtension(txt_filePath.Text.Trim())).ToLower() == ".xls")
            {
                strConn = "Provider=Microsoft.Jet.OLEDB.4.0;" + "data source=" + txt_filePath.Text.Trim() + ";Extended Properties=Excel 5.0;Persist Security Info=False";
                //sConnString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + txtFilePath.Text.Trim() + ";Extended Properties=\"Excel 8.0;HDR=" + strHead + ";IMEX=1\"";  
            }
            string sheetName = cmb_sheetName.Text.Trim();
            string strExcel = "select * from [" + sheetName + "]";
            try
            {
                ole = new OleDbConnection(strConn);
                ole.Open();
                da = new OleDbDataAdapter(strExcel, ole);
                dt_view = new DataTable();
                da.Fill(dt_view);
                if (dt_view.Rows.Count > 0)
                {
                    #region 处理
                    //*********************************************************************//
                    //因为生成Excel的时候第一行是标题，所以要做如下操作：  
                    //1.修改DataGridView列头的名字，  
                    //2.数据列表中删除第一行  
                    //for (int i = 0; i < dt.Columns.Count; i++)
                    //{
                    //    //dgvdata.Columns[i].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;  
                    //    //dgvdata.Columns[i].Name = dt.Columns[i].ColumnName;  
                    //    dgv_pvw.Columns[i].HeaderCell.Value = dt.Rows[0][i].ToString();//c# winform 用代码修改DataGridView列头的名字，设置列名,修改列名  
                    //}
                    ////DataGridView删除行  
                    //dgv_pvw.Rows.Remove(dgv_pvw.Rows[0]);//删除第一行  
                    //                                     //dgvdata.Rows.Remove(dgvdata.CurrentRow);//删除当前光标所在行  
                    //                                     //dgvdata.Rows.Remove(dgvdata.Rows[dgvdata.Rows.Count - 1]);//删除最后一行  
                    //                                     //dgvdata.Rows.Clear();//删除所有行  
                    //*********************************************************************//
                    List<string> columnStr = new List<string>();
                    //List<XMLTableColumnInfo> tcL_xml = LoadXMLTableColumn();
                    for (int i = 0; i < tcL_xml.Count; i++)
                    {
                        columnStr.Add(tcL_xml[i].Title);
                    }
                    dt_view = dt_view.DefaultView.ToTable(false, columnStr.ToArray());
                    //*********************************************************************//
                    this.dgv_pvw.DataSource = dt_view;
                    dgv_pvw.AllowUserToAddRows = false;//最后的空白行
                    dgv_pvw.RowHeadersVisible = false;//最后的空白行
                    dgv_pvw.AllowUserToResizeRows = false;//禁止用户拖动行高
                    #endregion
                }
                else
                {
                    this.dgv_pvw.DataSource = dt_view;
                    MessageBox.Show("无数据");
                }
                ole.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                if (ole != null)
                    ole.Close();
            }
            #endregion
        }

        /// <summary>
        /// Excel的Sheet下拉
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void cmb_sheetName_SelectedValueChanged(object sender, EventArgs e)
        {
            LoadExcel();
        }

        private void LoadXmlTable()
        {
            cmb_modelList.ValueMember = "Value";
            cmb_modelList.DisplayMember = "Text";

            string xmlPath = Application.StartupPath + @"\ExcelModelConfig\TableList.xml";
            XmlDocument xmlDoc = new XmlDocument();
            if (!File.Exists(xmlPath))
            {
                throw new Exception("配置文件不存在，路径：" + xmlPath);
            }
            xmlDoc.Load(xmlPath);
            XmlNodeList rootNode = xmlDoc.SelectSingleNode("root").ChildNodes;
            List<ComboboxInfo> cmbInfoL = new List<ComboboxInfo>();
            foreach (XmlNode node in rootNode)
            {
                string name = ((XmlElement)node).GetElementsByTagName("name")[0].InnerText.ToString();
                string title = ((XmlElement)node).GetElementsByTagName("title")[0].InnerText.ToString();
                ComboboxInfo ci = new ComboboxInfo();
                ci.Text = title;
                ci.Value = name;
                cmbInfoL.Add(ci);
            }
            cmb_modelList.DataSource = cmbInfoL;
            if (!string.IsNullOrEmpty(cmb_modelList.SelectedValue.ToString()))
            {
                table_import = cmb_modelList.SelectedValue.ToString();
                LoadXMLTableColumn(table_import);
            }
        }



        /// <summary>
        /// 加载表结构（XML）
        /// </summary>
        /// <returns></returns>
        private List<XMLTableColumnInfo> LoadXMLTableColumn(string table_import)
        {
            //string modelName = cmb_modelList.SelectedValue.ToString();
            string xmlPath = Application.StartupPath + @"\ExcelModelConfig\TableConfig\" + table_import + ".xml";
            XmlDocument xmlDoc = new XmlDocument();
            if (!File.Exists(xmlPath))
            {
                throw new Exception("配置文件不存在，路径：" + xmlPath);
            }
            xmlDoc.Load(xmlPath);
            XmlNodeList rootNode = xmlDoc.SelectSingleNode("root").ChildNodes;
            tcL_xml = new List<XMLTableColumnInfo>();
            foreach (XmlNode node in rootNode)
            {
                XmlElement columnNode = (XmlElement)node;
                string name = columnNode.GetElementsByTagName("name")[0].InnerText.ToString();
                string title = columnNode.GetElementsByTagName("title")[0].InnerText.ToString();
                string type = columnNode.GetElementsByTagName("type")[0].InnerText.ToString();
                int key = 0;
                if (columnNode.GetElementsByTagName("key")[0] != null)
                {
                    key = Convert.ToInt32(columnNode.GetElementsByTagName("key")[0].InnerText);
                }
                XMLTableColumnInfo xml_tc = new XMLTableColumnInfo();
                xml_tc.Name = name;
                xml_tc.Title = title;
                xml_tc.Type = type;
                xml_tc.Key = key;
                tcL_xml.Add(xml_tc);
            }
            return tcL_xml;
        }

        private void btn_Import_Click(object sender, EventArgs e)
        {
            List<string> columnStr = new List<string>();
            for (int i = 0; i < tcL_xml.Count; i++)
            {
                columnStr.Add(tcL_xml[i].Name);
            }

            StringBuilder import_sql = new StringBuilder();//导入的SQL语句
            import_sql.AppendFormat(" INSERT INTO {0} ( ", table_import);
            import_sql.AppendFormat(" {0} ", string.Join(",", columnStr.ToArray()));
            import_sql.Append(" ) VALUES ( ");
            import_sql.AppendFormat(" @{0} ", string.Join(",@", columnStr.ToArray()));
            import_sql.Append(" ) ");
            List<SqlParameter[]> paramList = new List<SqlParameter[]>();
            foreach (DataRow dr in dt_view.Rows)
            {
                List<SqlParameter> paramters = new List<SqlParameter>();
                for (int i = 0; i < tcL_xml.Count; i++)
                {
                    string title = tcL_xml[i].Title;
                    string name = tcL_xml[i].Name;
                    paramters.Add(new SqlParameter("@" + name, dr[title]));
                }
                paramList.Add(paramters.ToArray());
            }
            SqlHelper db = new SqlHelper();
            string ConnectionString = @"server = BANANA\MSSQLSERVER2014; uid = sa; pwd = Sa123456; database = wzjy";
            int r = db.ExecuteSqlStr(import_sql.ToString(), paramList, ConnectionString);
            if (r > 0)
            {
                MessageBox.Show("成功导入数据 " + r + " 条");
            }
            else
            {
                MessageBox.Show("导入失败");
            }
        }

        #region
        /// <summary>
        /// 下拉框
        /// </summary>
        private class ComboboxInfo
        {
            public string Text { get; set; }
            public string Value { get; set; }
        }
        /// <summary>
        /// 表字段信息（XML)
        /// </summary>
        private class XMLTableColumnInfo
        {
            public string Name { get; set; }//字段名称
            public string Title { get; set; }//字段中文名
            public string Type { get; set; }//字段数据类型
            public int Key { get; set; }//1 主键 0 否 => 数据唯一标识
        }
        #endregion

    }
}
