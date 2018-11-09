using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Text.RegularExpressions;
using System.Threading; 

using Autodesk.Navisworks.Api;
using Autodesk.Navisworks.Api.Plugins;
using Autodesk.Navisworks.Api.ComApi;

namespace ClassLibraryNavisworksROOMDFS
{
    public partial class DFSControl : UserControl
    {
        public Document oDoc;
        DataTable dtDFS = new DataTable();

        DataTable dtFROM = new DataTable();
        DataTable dtTO = new DataTable();

        DataTable dtPaths = new DataTable();

        private DataSet dSet = new DataSet();

        List<cNode> listNode = new List<cNode>();
        List<cNodePath> listPath = new List<cNodePath>();
        List<cNodePathRecord> listPathDFS = new List<cNodePathRecord>();

        int iPathNo = 0; //路径记录号

        public DFSControl()
        {
            InitializeComponent();

            dtDFS.Columns.Add("ID", System.Type.GetType("System.Decimal"));//1
            dtDFS.Columns.Add("起始空间", System.Type.GetType("System.String"));//1
            dtDFS.Columns.Add("结束空间", System.Type.GetType("System.String"));//1
            dtDFS.Columns.Add("途径空间数", System.Type.GetType("System.Decimal"));//1
            dtDFS.Columns.Add("途径门数", System.Type.GetType("System.Decimal"));//1
            dtDFS.Columns.Add("路径描述", System.Type.GetType("System.String"));//1

            dataGridViewDFS.DataSource = dtDFS;
            for (int i = 0; i < dataGridViewDFS.ColumnCount; i++)
            {
                dataGridViewDFS.Columns[i].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            }


            dtPaths.Columns.Add("类", System.Type.GetType("System.String"));//1
            dtPaths.Columns.Add("描述", System.Type.GetType("System.String"));//1
            dtPaths.Columns.Add("EID", System.Type.GetType("System.String"));//1
            dataGridViewPaths.DataSource = dtPaths;
            for (int i = 0; i < dataGridViewPaths.ColumnCount; i++)
            {
                dataGridViewPaths.Columns[i].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            }
            dataGridViewPaths.Columns[2].Visible = false;


            dtFROM.Columns.Add("ID", System.Type.GetType("System.Decimal"));//1
            dtFROM.Columns.Add("空间名称", System.Type.GetType("System.String"));//1
            dtFROM.Columns.Add("空间ID", System.Type.GetType("System.String"));//1


            comboBoxFrom.DataSource = dtFROM;
            comboBoxFrom.DisplayMember = "空间名称";
            comboBoxFrom.ValueMember = "空间ID";

            dtTO.Columns.Add("ID", System.Type.GetType("System.Decimal"));//1
            dtTO.Columns.Add("空间名称", System.Type.GetType("System.String"));//1
            dtTO.Columns.Add("空间ID", System.Type.GetType("System.String"));//1

            comboBoxTo.DataSource = dtTO;
            comboBoxTo.DisplayMember = "空间名称";
            comboBoxTo.ValueMember = "空间ID";



            oDoc = Autodesk.Navisworks.Api.Application.ActiveDocument;
        }

        private void btnShow_Click(object sender, EventArgs e)
        {
            

            VariantData oData;
            SearchCondition oSearchCondition;
            ModelItemCollection items;


            string pattern = @"\{[^\{^\}]*\}";
            string pattern1 = @"\[[^\[^\]]*\]";
            Regex regex = new Regex(pattern, RegexOptions.IgnoreCase);
            Regex regex1 = new Regex(pattern1, RegexOptions.IgnoreCase);


            //清除表达
            clearAppearance();

            int i;
            string sT = "";

            for (i = 0; i < dataGridViewDFS.SelectedRows.Count; i++)
            {
                sT = dataGridViewDFS.SelectedRows[i].Cells[5].Value.ToString();
                //MessageBox.Show(sT);

                //分离房间

                MatchCollection matches = regex.Matches(sT);
                //StringBuilder sb = new StringBuilder();//存放匹配结果
                foreach (Match match in matches)
                {
                    //Thread.Sleep(1000);
                    string value = match.Value.Trim('{', '}');
                    

                    string[] sArray=value.Split(':');
                    if(sArray.Length!=2)
                        continue;
                    //MessageBox.Show(value + "sArray[0]=" + sArray[0] + "sArray[1]="+sArray[1]);
                    //找到房间
                    Search search = new Search();
                    search.Selection.SelectAll();

                    oData = VariantData.FromDisplayString(sArray[0]);

                    oSearchCondition = SearchCondition.HasPropertyByDisplayName("元素 ID", "值");
                    oSearchCondition = oSearchCondition.EqualValue(oData);
                    search.SearchConditions.Add(oSearchCondition);
                    items = search.FindAll(oDoc, false);

                    
                    switch(sArray[1])
                    {
                        case "0":
                            oDoc.Models.OverridePermanentColor(items, Autodesk.Navisworks.Api.Color.Red);
                            oDoc.Models.OverridePermanentTransparency(items, 0.5);
                            break;
                        case "1":
                            oDoc.Models.OverridePermanentColor(items, Autodesk.Navisworks.Api.Color.Blue);
                            oDoc.Models.OverridePermanentTransparency(items, 0.6);
                            break;
                        case "2":
                            oDoc.Models.OverridePermanentColor(items, Autodesk.Navisworks.Api.Color.Green);
                            oDoc.Models.OverridePermanentTransparency(items, 0.9);
                            break;
                    }
                    

                }

                //分离门
                matches = regex1.Matches(sT);
                //StringBuilder sb1 = new StringBuilder();//存放匹配结果
                foreach (Match match in matches)
                {
                    
                    string value = match.Value.Trim('[', ']');
                    
                    string[] sArray1 = value.Split(',');


                    foreach (string s in sArray1)
                    {
                        //MessageBox.Show(value + "sArray[0]=" + sArray[0] + "sArray[1]="+sArray[1]);
                        //找到
                        Search search = new Search();
                        search.Selection.SelectAll();

                        //MessageBox.Show(value);

                        oData = VariantData.FromDisplayString(s);

                        oSearchCondition = SearchCondition.HasPropertyByDisplayName("元素 ID", "值");
                        oSearchCondition = oSearchCondition.EqualValue(oData);
                        search.SearchConditions.Add(oSearchCondition);
                        items = search.FindAll(oDoc, false);
                        oDoc.Models.OverridePermanentColor(items, Autodesk.Navisworks.Api.Color.Red);
                        oDoc.Models.OverridePermanentTransparency(items, 0);
                    }

                }

            }



        }

        private void clearAppearance() //清除所有表达
        {
            Search search = new Search();
            search.Selection.SelectAll();

            //oDoc.Models.ResetAllPermanentMaterials();

            VariantData oData = VariantData.FromDisplayString("门");

            SearchCondition oSearchCondition = SearchCondition.HasPropertyByDisplayName("项目", "类型");
            oSearchCondition = oSearchCondition.EqualValue(oData);
            search.SearchConditions.Add(oSearchCondition);
            ModelItemCollection items = search.FindAll(oDoc, false);
            oDoc.Models.OverridePermanentColor(items, Autodesk.Navisworks.Api.Color.Blue);
            oDoc.Models.OverridePermanentTransparency(items, 0.5);

            Search search1 = new Search();
            search1.Selection.SelectAll();
            oData = VariantData.FromDisplayString("房间");

            oSearchCondition = SearchCondition.HasPropertyByDisplayName("项目", "类型");
            oSearchCondition = oSearchCondition.EqualValue(oData);
            search1.SearchConditions.Add(oSearchCondition);
            items = search1.FindAll(oDoc, false);
            oDoc.Models.OverridePermanentColor(items, Autodesk.Navisworks.Api.Color.White);
            oDoc.Models.OverridePermanentTransparency(items, 1);
            //oDoc.CurrentSelection.CopyFrom(items);
        }

        private void btnR_Click(object sender, EventArgs e)
        {
            int i=0;
            try
            {
                OpenFileDialog openFileDialogOutput = new OpenFileDialog();
                openFileDialogOutput.Filter = "xml files(*.xml)|*.xml";//excel files(*.xls)|*.xls|All files(*.*)|*.*
                openFileDialogOutput.FilterIndex = 0;
                openFileDialogOutput.RestoreDirectory = true;


                if (openFileDialogOutput.ShowDialog() != DialogResult.OK) return;

                dSet.Tables.Clear();
                dSet.ReadXml(openFileDialogOutput.FileName.ToString());
                initDataList(); dtDFS.Clear();
                InitDataGridViewDFS("","");

                //初始化过滤块
                object[] oTemp = new object[3];
                dtFROM.Clear();dtTO.Clear();

                //comboBoxFrom.Items.Clear();
                //comboBoxTo.Items.Clear();

                oTemp[0]=0;oTemp[1]="ALL";oTemp[2]="";
                dtFROM.Rows.Add(oTemp);
                dtTO.Rows.Add(oTemp);

                i=0;
                var qfrom =from cNode cnode in listNode
                           where cnode.iROOMCLASS==2
                           select cnode;

                foreach(cNode cnode in qfrom)
                {
                    i++;
                    oTemp[0]=i;oTemp[1]=cnode.sROOMName;oTemp[2]=cnode.iROOMID;
                    dtFROM.Rows.Add(oTemp);
                }




                i = 0;
                var qto = from cNode cnode in listNode
                            where cnode.iROOMCLASS == 0
                            select cnode;

                foreach (cNode cnode in qto)
                {
                    i++;
                    oTemp[0] = i; oTemp[1] = cnode.sROOMName; oTemp[2] = cnode.iROOMID;
                    dtTO.Rows.Add(oTemp);
                }


                
            }
            catch (System.Exception err)
            {
                MessageBox.Show("读取文件失败" + err.Message);
            }
        }

        private void InitDataGridViewDFS(string sFromFilter, string sToFilter)
        {
            int i,j;

            object[] oTempDFS = new object[6];
            cNodePathRecord cnodeprF=null,cnodeprT;
            int iMax;


            int ipaths=0;
            i = 1;
            while(true)
            {
                //MessageBox.Show(ipaths.ToString());
                ipaths++;
                var qPathDFS = from cPathDFS in listPathDFS
                               where cPathDFS.iPathNO==ipaths
                               orderby cPathDFS.iPathNO , cPathDFS.iPathNO ascending
                               select cPathDFS;
                
                if(qPathDFS.Count()<1) //没有路径，中止
                    break;

                oTempDFS[0] = ipaths;

                //路径描述
                j=0;iMax=0;oTempDFS[5]="";
                foreach (cNodePathRecord cnodepr in qPathDFS)
                {
                    if(j==0) //第一点
                    {
                        cnodeprF=cnodepr;
                        oTempDFS[5] = cnodepr.sROOMName + "{" + cnodepr.iROOMID + ":" + cnodepr .iROOMCLASS.ToString()+ "}";
                        oTempDFS[1] = oTempDFS[5];
                        oTempDFS[2] = oTempDFS[1];
                        j++;
                    }
                    else //中间点，取路径
                    {

                            
                        var qPaths = from cPaths in listPath
                                        where cPaths.nToROOM.iROOMID ==cnodeprF.iROOMID && cPaths.nFromROOM.iROOMID ==cnodepr.iROOMID
                                        select cPaths;

                        //MessageBox.Show(cnodeprF.iROOMID + "-" + qPaths.Count().ToString() + "-" + cnodepr.iROOMID.ToString());

                        if(qPaths.Count()<1)
                        {
                            cnodeprF=cnodepr;
                            continue;
                        }

                        foreach(cNodePath cpath in qPaths)
                        {
                            iMax+=cpath.iValue;
                            cnodeprF=cnodepr;
                            oTempDFS[5] += "->["+cpath.sDoorID+"]->" + cnodepr.sROOMName + "{" + cnodepr.iROOMID + ":" + cnodepr .iROOMCLASS.ToString()+ "}";
                            oTempDFS[2] = cnodepr.sROOMName + "{" + cnodepr.iROOMID + ":" + cnodepr.iROOMCLASS.ToString() + "}";
                            j++;
                        }

                    }
                        
                }

                oTempDFS[3] = j;oTempDFS[4] = iMax;

                if (sFromFilter == "" && sToFilter == "") //没有查询条件
                {                    
                    dtDFS.Rows.Add(oTempDFS);
                    i++;
                }
                else //有约束条件
                {
                    if (sFromFilter != "" && sToFilter != "") //均有约束
                    {
                        if (oTempDFS[1].ToString().Contains(sFromFilter) && oTempDFS[2].ToString().Contains(sToFilter))
                        {
                            dtDFS.Rows.Add(oTempDFS);
                            i++;
                        }
                    }
                    else //单一约束
                    {
                        if (sFromFilter != "")
                        {
                            if (oTempDFS[1].ToString().Contains(sFromFilter))
                            {
                                dtDFS.Rows.Add(oTempDFS);
                                i++;
                            }
                        }
                        else
                        {
                            if (oTempDFS[2].ToString().Contains(sToFilter))
                            {
                                dtDFS.Rows.Add(oTempDFS);
                                i++;
                            }
                        }
                    }
                    
                    
                }
            }

            toolStripStatusLabelS.Text = "共有"+dataGridViewDFS.RowCount.ToString()+"条路径";
            dtPaths.Rows.Clear();

        }



        private void initDataList() //数据还原，太懒了，简单点吧
        {
            int i;

            listNode.Clear();
            listPath.Clear();
            listPathDFS.Clear();


            for (i = 0; i < dSet.Tables["NODE"].Rows.Count; i++)
            {
                cNode cnode = new cNode();
                cnode.iROOMID = int.Parse(dSet.Tables["NODE"].Rows[i][0].ToString());
                cnode.sROOMName = dSet.Tables["NODE"].Rows[i][1].ToString();
                cnode.iROOMCLASS = int.Parse(dSet.Tables["NODE"].Rows[i][2].ToString());
                //cnode.sROOMName = dSet.Tables["NODE"].Rows[i][3].ToString();
                listNode.Add(cnode);
            }

            for (i = 0; i < dSet.Tables["PATH"].Rows.Count; i++)
            {
                cNodePath cnodepath = new cNodePath();
                
                //找到起始节点
                var qNodes = from cnode in listNode
                             where cnode.iROOMID == int.Parse(dSet.Tables["PATH"].Rows[i][0].ToString())
                             select cnode;
                if (qNodes.Count() < 1)
                    continue;
                foreach (cNode cnode1 in qNodes)
                {
                    cnodepath.nFromROOM = cnode1;
                    break;
                }

                //找到终点
                var qNodes1 = from cnode in listNode
                             where cnode.iROOMID == int.Parse(dSet.Tables["PATH"].Rows[i][2].ToString())
                             select cnode;
                if (qNodes1.Count() < 1)
                    continue;
                foreach (cNode cnode1 in qNodes1)
                {
                    cnodepath.nToROOM = cnode1;
                    break;
                }
                cnodepath.iValue = int.Parse(dSet.Tables["PATH"].Rows[i][4].ToString());
                cnodepath.sDoorID = dSet.Tables["PATH"].Rows[i][5].ToString();
                cnodepath.sDoorName = dSet.Tables["PATH"].Rows[i][6].ToString();


                listPath.Add(cnodepath);
            }
            

            for (i = 0; i < dSet.Tables["DFS"].Rows.Count; i++)
            {
                cNodePathRecord cnodepathr = new cNodePathRecord();

                //找到节点
                cnodepathr.iPathNO = int.Parse(dSet.Tables["DFS"].Rows[i][0].ToString());
                cnodepathr.iNodeNO = int.Parse(dSet.Tables["DFS"].Rows[i][1].ToString());
                cnodepathr.iROOMID = int.Parse(dSet.Tables["DFS"].Rows[i][2].ToString());
                cnodepathr.sROOMName = dSet.Tables["DFS"].Rows[i][3].ToString();
                cnodepathr.iROOMCLASS = int.Parse(dSet.Tables["DFS"].Rows[i][7].ToString());
                listPathDFS.Add(cnodepathr);
            }

            var pathMax = (from s in listPathDFS
                          select s.iPathNO)
                      .Max();
            iPathNo = pathMax;
            //MessageBox.Show(iPathNo.ToString());

        }

        private void btnClear_Click(object sender, EventArgs e)
        {
            clearAppearance();
        }

        //路径描述
        private void dataGridViewDFS_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridViewDFS.RowCount < 1)
                return;

            dtPaths.Rows.Clear(); 

            if (dataGridViewDFS.SelectedRows.Count < 1)
            {
                dtPaths.Rows.Clear();
                return;
            }

            var qPathDFS = from cPathDFS in listPathDFS
                           where cPathDFS.iPathNO == int.Parse(dataGridViewDFS.SelectedRows[0].Cells[0].Value.ToString())
                           orderby cPathDFS.iPathNO, cPathDFS.iPathNO ascending
                           select cPathDFS;

            if (qPathDFS.Count() < 1) //没有路径，中止
            {
                dtPaths.Rows.Clear();
                return;
            }


            int j = 0;
            cNodePathRecord cnodeprF = null;
            object[] oT = new object[3];
            foreach (cNodePathRecord cnodepr in qPathDFS)
            {
                if (j == 0) //第一点
                {
                    cnodeprF = cnodepr;
                    oT[0] = "空间";
                    oT[1] = cnodepr.sROOMName;
                    oT[2] = cnodepr.iROOMID.ToString() + ":" + cnodepr.iROOMCLASS;

                    dtPaths.Rows.Add(oT);
                    j++;
                }
                else //中间点，取路径
                {
                    var qPaths = from cPaths in listPath
                                 where cPaths.nToROOM.iROOMID == cnodeprF.iROOMID && cPaths.nFromROOM.iROOMID == cnodepr.iROOMID
                                 select cPaths;

                    if (qPaths.Count() < 1)
                    {
                        cnodeprF = cnodepr;
                        continue;
                    }

                    foreach (cNodePath cpath in qPaths)
                    {
                        cnodeprF = cnodepr;

                        oT[0] = "门";
                        oT[1] = cpath.sDoorName;
                        oT[2] = cpath.sDoorID;
                        dtPaths.Rows.Add(oT);

                        oT[0] = "空间";
                        oT[1] = cnodepr.sROOMName;
                        oT[2] = cnodepr.iROOMID.ToString()+":" + cnodepr.iROOMCLASS;;
                        dtPaths.Rows.Add(oT);
                        //textBoxPaths.Text += "\r\n->[" + cpath.sDoorName+"]->\r\n" + cnodepr.sROOMName + "}";

                        j++;
                    }

                }

            }

            //textBoxPaths.Text += "\r\n\r\n经：\r\n" + dataGridViewDFS.SelectedRows[0].Cells[3].Value.ToString() + "个空间\r\n" + dataGridViewDFS.SelectedRows[0].Cells[4].Value.ToString() + "扇门";
            toolStripStatusLabelS.Text = "共有" + dataGridViewDFS.RowCount.ToString() + "条路径，所选路径经：" + dataGridViewDFS.SelectedRows[0].Cells[3].Value.ToString() + "个空间、" + dataGridViewDFS.SelectedRows[0].Cells[4].Value.ToString() + "扇门";




        }

        //过滤
        private void btnFilter_Click(object sender, EventArgs e)
        {
            dtDFS.Clear();
            if (listNode.Count < 1)
                return;
            InitDataGridViewDFS(comboBoxFrom.SelectedValue.ToString(),comboBoxTo.SelectedValue.ToString());
        }

        private void dataGridViewPaths_DoubleClick(object sender, EventArgs e)
        {
            if (dataGridViewPaths.RowCount < 1)
                return;

            if (dataGridViewPaths.SelectedRows.Count < 1)
                return;



            VariantData oData;
            SearchCondition oSearchCondition;
            ModelItemCollection items;



            //清除表达
            //clearAppearance();

            int i;
            oDoc.CurrentSelection.Clear();

            for (i = 0; i < dataGridViewPaths.SelectedRows.Count; i++)
            {
                if (dataGridViewPaths.SelectedRows[i].Cells[0].Value.ToString() == "空间")
                {
                    string[] sArray = dataGridViewPaths.SelectedRows[i].Cells[2].Value.ToString().Split(':');
                    if (sArray.Length != 2)
                        continue;
                    //MessageBox.Show(value + "sArray[0]=" + sArray[0] + "sArray[1]="+sArray[1]);
                    //找到房间
                    Search search = new Search();
                    search.Selection.SelectAll();

                    oData = VariantData.FromDisplayString(sArray[0]);

                    oSearchCondition = SearchCondition.HasPropertyByDisplayName("元素 ID", "值");
                    oSearchCondition = oSearchCondition.EqualValue(oData);
                    search.SearchConditions.Add(oSearchCondition);
                    items = search.FindAll(oDoc, false);


                    oDoc.CurrentSelection.CopyFrom(items);

                    /*

                    switch (sArray[1])
                    {
                        case "0":
                            oDoc.Models.OverridePermanentColor(items, Autodesk.Navisworks.Api.Color.Red);
                            oDoc.Models.OverridePermanentTransparency(items, 0.5);
                            break;
                        case "1":
                            oDoc.Models.OverridePermanentColor(items, Autodesk.Navisworks.Api.Color.Blue);
                            oDoc.Models.OverridePermanentTransparency(items, 0.6);
                            break;
                        case "2":
                            oDoc.Models.OverridePermanentColor(items, Autodesk.Navisworks.Api.Color.Green);
                            oDoc.Models.OverridePermanentTransparency(items, 0.9);
                            break;
                    }
                     * */

                }
                else
                {

                    string[] sArray1 = dataGridViewPaths.SelectedRows[i].Cells[2].Value.ToString().Split(',');


                    foreach (string s in sArray1)
                    {
                        //MessageBox.Show(value + "sArray[0]=" + sArray[0] + "sArray[1]="+sArray[1]);
                        //找到
                        Search search = new Search();
                        search.Selection.SelectAll();

                        //MessageBox.Show(value);

                        oData = VariantData.FromDisplayString(s);

                        oSearchCondition = SearchCondition.HasPropertyByDisplayName("元素 ID", "值");
                        oSearchCondition = oSearchCondition.EqualValue(oData);
                        search.SearchConditions.Add(oSearchCondition);
                        items = search.FindAll(oDoc, false);

                        oDoc.CurrentSelection.CopyFrom(items);

                        /*
                        oDoc.Models.OverridePermanentColor(items, Autodesk.Navisworks.Api.Color.Red);
                        oDoc.Models.OverridePermanentTransparency(items, 0);
                        */
                    }

                }

            }

            if ((Control.ModifierKeys & Keys.Control) == Keys.Control)
            {
                ComApiBridge.State.ZoomInCurViewOnCurSel();
            }

            //ComApiBridge.State.ZoomInCurViewOnCurSel(); 


        }

        private void dataGridViewDFS_KeyPress(object sender, KeyPressEventArgs e)
        {

        }

        public void DataGridViewEnableCopy(DataGridView p_Data)
        {

            Clipboard.SetData(DataFormats.Text, p_Data.GetClipboardContent());

        }

        private void dataGridViewDFS_KeyDown(object sender, KeyEventArgs e)
        {
            if (Control.ModifierKeys == Keys.Control && e.KeyCode == Keys.C)
            {

                if (sender != null && sender.GetType() == typeof(DataGridView))

                    DataGridViewEnableCopy((DataGridView)sender);

            }
        }
    }

    //节点结构,房间
    public class cNode
    {
        public int iROOMID = 0; //节点ID
        public string sROOMName = "";

        //房间类型,0=保护空间，1=通道，2=外部空间   -1=无效
        public int iROOMCLASS = -1;

        public Boolean bFlag = false;

    }

    //邻接表存储结构
    public class cNodePath
    {
        public cNode nFromROOM; //主节点

        public cNode nToROOM; //到节点

        public int iValue = 0; //取值，门数

        //public ArrayList alDoor = new ArrayList(); //门列表
        //public ArrayList alDoorN = new ArrayList(); //门列表
        public string sDoorID = "";
        public string sDoorName = "";

        //临时标记
        public Boolean bFlag = false;
    }

    //路径结果结构,房间
    public class cNodePathRecord
    {
        public int iROOMID = 0; //节点ID
        public string sROOMName = "";

        //房间类型,0=保护空间，1=通道，2=外部空间   -1=无效
        public int iROOMCLASS = -1;

        //public int iValue = 0; //取值，门数

        //public string sDoorID = "";
        //public string sDoorName = "";

        public int iPathNO = -1;
        public int iNodeNO = -1;

    }
}
