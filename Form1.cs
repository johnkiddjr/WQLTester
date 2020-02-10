using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using Microsoft.Management;
using Microsoft.Management.Infrastructure;

namespace WQLTester
{
    public partial class frmMain : Form
    {
        public frmMain()
        {
            InitializeComponent();
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void btnClear_Click(object sender, EventArgs e)
        {
            dgvResults.DataSource = null;
        }

        private void btnTest_Click(object sender, EventArgs e)
        {
            //try to connect to the server using built in credentials
            try
            {
                CimSession wmisession = CimSession.Create(txtSiteServer.Text);
                
                if (wmisession.TestConnection())
                {
                    MessageBox.Show("Test Succeeded!");
                }
                else
                {
                    MessageBox.Show("Unable to connect to server/namespace!");
                }
            }
            catch
            {
                MessageBox.Show("Unable to connect to server/namespace!");
            }
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            try
            {
                CimSession wmisession = CimSession.Create(txtSiteServer.Text);
                IEnumerable<CimInstance> results = wmisession.QueryInstances(txtWmiClass.Text, "WQL", txtQuery.Text);
                if (results.Count() > 0)
                {
                    //create a datatable
                    DataTable dt = new DataTable();

                    //create the columns we care about
                    CimInstance firstIntance = results.First();
                    foreach (CimProperty _property in firstIntance.CimInstanceProperties)
                    {
                        if (_property.Value != null)
                        {
                            dt.Columns.Add(_property.Name, typeof(string));
                        }
                    }

                    //enter data into datatable
                    foreach (CimInstance _instance in results)
                    {
                        DataRow _newrow = dt.NewRow();
                        foreach (DataColumn _column in dt.Columns)
                        {
                            _newrow[_column] = _instance.CimInstanceProperties[_column.ColumnName].Value;
                        }

                        dt.Rows.Add(_newrow);
                    }

                    //mount datatable onto datagrid
                    dgvResults.DataSource = dt;
                }
            }
            catch
            {
                MessageBox.Show("Query Failure, check network connections!");
            }
        }
    }
}
