//==============================================================================
// TITLE: Form1.cs - Ver2.00(VCDotNetRcwSample)
//
// CONTENTS:
//
// (c) Copyright 2004-2010 Takebishi Electric Sales Corporation
// ALL RIGHTS RESERVED.
//
// DISCLAIMER:
//  This code is provided by the Takebishi Corporation solely to assist 
//  in understanding and use of the appropriate OPC Specification(s).
//  This code is provided as-is and without warranty or support of any sort.
//
// MODIFICATION LOG:
//
// Version		Date		By			Notes
// --------		--------	--------	--------
//' 1.00		2004/02/27	Mike		First release.
//' 2.00		2010/02/05	ykishimo	Fixed ini file path for Windows Vista and Windows7.
//' 2.10		2012/01/17	ykishimo	Support Shutdown event and Group Active/InActive state setting.
//'==============================================================================

using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.IO;
using System.Runtime.InteropServices;
using OpcRcw.Da;
using System.Threading;	// 06/09/20 for VS2005
using System.Collections.Generic;

namespace VCDotNetRcwSample
{
	/// <summary>
	/// Form1
	/// </summary>
	public class Form1 : System.Windows.Forms.Form
	{
		// This delegate enables asynchronous calls for setting
		// the text property on a TextBox control.
		delegate void SetValueCallback(int i);      // 06/09/20 for VS2005
		delegate void SetTimeCallback(int i);       // 06/09/20 for VS2005
		delegate void SetQualityCallback(int i);    // 06/09/20 for VS2005

		// This thread is used to demonstrate both thread-safe and
		// unsafe ways to call a Windows Forms control.
		private Thread threadCopy = null;           // 06/09/20 for VS2005

		/// <summary>
		/// Designer variable
		/// </summary>
		private System.ComponentModel.Container components = null;

		private enum DEF_COPY_DIR
		{
			DISP_TO_MEM = 0,	// display to memory
			MEM_TO_DISP,		// memory to display
		}

		[StructLayoutAttribute(LayoutKind.Explicit)]
		private struct JointInt
		{
			// long (int * 2 and long * 1 share the same memory (64bits)
			[FieldOffsetAttribute(0)] public long lJoint;

			// int (int * 2 and long * 1 share the same memory (64bits)
			[FieldOffsetAttribute(0)] public int iLo;
			[FieldOffsetAttribute(4)] public int iHi;
/*
			public JointInt()
			{
				lJoint = 0;
				iLo = 0;
				iHi = 0;
			}
*/
		}

		private static readonly int ITEMMAX				= 10;
		private static readonly string OLGA_OPC_INI		= "OLGA_OPC_Component.ini";

		private static readonly int VAL_CTRL_SPACE = 6;
		private static readonly int VAL_CTRL_TOP = 230;
		private static readonly int VAL_CTRL_LEFT = 16;
		private static readonly int VAL_ITEMNAME_HEIGHT = 24;
		private static readonly int VAL_ITEMNAME_WIDTH = 200;
		private static readonly int VAL_VALUE_HEIGHT = VAL_ITEMNAME_HEIGHT;
		private static readonly int VAL_VALUE_WIDTH = 100;
		private static readonly int VAL_TIME_HEIGHT = VAL_ITEMNAME_HEIGHT;
		private static readonly int VAL_TIME_WIDTH = 100;
		private static readonly int VAL_QUALITY_HEIGHT = VAL_ITEMNAME_HEIGHT;
		private static readonly int VAL_QUALITY_WIDTH = 40;

		private COPCServer OPCSvr;

		private string[] sItemName = new string[ITEMMAX];
		private int[] cH = new int[ITEMMAX];
		private int[] sH = new int[ITEMMAX];
		private object[] oVal = new object[ITEMMAX];
		private OpcRcw.Da.FILETIME[] dTime = new OpcRcw.Da.FILETIME[ITEMMAX];
		private short[] wQuality = new short[ITEMMAX];

		private System.Windows.Forms.TextBox[] txtItemName	= new System.Windows.Forms.TextBox[ITEMMAX];
		private System.Windows.Forms.TextBox[] txtValue		= new System.Windows.Forms.TextBox[ITEMMAX];
		private System.Windows.Forms.TextBox[] txtTime		= new System.Windows.Forms.TextBox[ITEMMAX];
		private System.Windows.Forms.TextBox[] txtQuality	= new System.Windows.Forms.TextBox[ITEMMAX];
		private DEF_OPCDA OpcdaVer;

		private Label label10;
		private String csInifilePath;
        private CheckBox chkGrpActive;
        private RadioButton rdoDA30;
        private RadioButton rdoDA20;
        private Label Label7;
		private String spFolder;

		public Form1()
		{
			//
			// For Windows form designer support
			//
			InitializeComponent();

			OPCSvr = new COPCServer();

			// Add events of COPCServer.
			OPCSvr.DataChange		+= new DataChangeHandler(OPCSvr_DataChange);
			OPCSvr.ReadComplete		+= new ReadCompleteHandler(OPCSvr_ReadComplete);
			OPCSvr.WriteComplete	+= new WriteCompleteHandler(OPCSvr_WriteComplete);
			OPCSvr.CancelComplete	+= new CancelCompleteHandler(OPCSvr_CancelComplete);
            OPCSvr.ShutDownRequestEvent += new ShutDownRequestHandler(OPCSvr_ShutDown);	// 2011/11/14 シャット?ウンイベントを受けれるようにする


			OpcdaVer = DEF_OPCDA.VER_NONE;
		}

		/// <summary>
		/// Execute dispose of used resources.
		/// </summary>
		protected override void Dispose( bool disposing )
		{
			if( disposing )
			{
				if (components != null) 
				{
					components.Dispose();
				}
			}
			base.Dispose( disposing );
		}

		#region Windows フォ?? デザイナで生成されたコ?ド 

		private System.Windows.Forms.TextBox txtUpdateRate;
		private System.Windows.Forms.TextBox txtNode;
        private System.Windows.Forms.TextBox txtGrp;
		private System.Windows.Forms.Button btnRefresh;
		private System.Windows.Forms.Button btnAsyncWrite;
		private System.Windows.Forms.Button btnAsyncRead;
		private System.Windows.Forms.Button btnMaxAge;
		private System.Windows.Forms.Button btnAdvise;
		private System.Windows.Forms.Button btnWrite;
		private System.Windows.Forms.Button btnRead;
		private System.Windows.Forms.Button btnConnect;
		private System.Windows.Forms.Label Label9;
        private System.Windows.Forms.Label Label8;
		private System.Windows.Forms.Label Label6;
		private System.Windows.Forms.Label Label5;
		private System.Windows.Forms.Label Label4;
		private System.Windows.Forms.Label Label3;
		private System.Windows.Forms.Label Label2;
		private System.Windows.Forms.Label Label1;
		private System.Windows.Forms.ComboBox cmbSvrName;

		/// <summary>
		/// デザイナ サ??トに必要なメ?ッドです。このメ?ッドの内容を
		/// コ?ド エディ?で変更しないでください。
		/// </summary>
		private void InitializeComponent()
		{
            this.txtUpdateRate = new System.Windows.Forms.TextBox();
            this.btnAsyncWrite = new System.Windows.Forms.Button();
            this.btnAsyncRead = new System.Windows.Forms.Button();
            this.btnMaxAge = new System.Windows.Forms.Button();
            this.btnAdvise = new System.Windows.Forms.Button();
            this.btnWrite = new System.Windows.Forms.Button();
            this.btnRead = new System.Windows.Forms.Button();
            this.btnConnect = new System.Windows.Forms.Button();
            this.Label9 = new System.Windows.Forms.Label();
            this.Label8 = new System.Windows.Forms.Label();
            this.Label6 = new System.Windows.Forms.Label();
            this.Label5 = new System.Windows.Forms.Label();
            this.Label4 = new System.Windows.Forms.Label();
            this.Label3 = new System.Windows.Forms.Label();
            this.Label2 = new System.Windows.Forms.Label();
            this.Label1 = new System.Windows.Forms.Label();
            this.txtNode = new System.Windows.Forms.TextBox();
            this.cmbSvrName = new System.Windows.Forms.ComboBox();
            this.btnRefresh = new System.Windows.Forms.Button();
            this.txtGrp = new System.Windows.Forms.TextBox();
            this.label10 = new System.Windows.Forms.Label();
            this.chkGrpActive = new System.Windows.Forms.CheckBox();
            this.rdoDA30 = new System.Windows.Forms.RadioButton();
            this.rdoDA20 = new System.Windows.Forms.RadioButton();
            this.Label7 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // txtUpdateRate
            // 
            this.txtUpdateRate.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtUpdateRate.Location = new System.Drawing.Point(114, 101);
            this.txtUpdateRate.Name = "txtUpdateRate";
            this.txtUpdateRate.Size = new System.Drawing.Size(108, 20);
            this.txtUpdateRate.TabIndex = 93;
            this.txtUpdateRate.Text = "1000";
            this.txtUpdateRate.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            // 
            // btnAsyncWrite
            // 
            this.btnAsyncWrite.BackColor = System.Drawing.SystemColors.Control;
            this.btnAsyncWrite.Cursor = System.Windows.Forms.Cursors.Default;
            this.btnAsyncWrite.Enabled = false;
            this.btnAsyncWrite.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnAsyncWrite.ForeColor = System.Drawing.SystemColors.ControlText;
            this.btnAsyncWrite.Location = new System.Drawing.Point(368, 190);
            this.btnAsyncWrite.Name = "btnAsyncWrite";
            this.btnAsyncWrite.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.btnAsyncWrite.Size = new System.Drawing.Size(125, 37);
            this.btnAsyncWrite.TabIndex = 102;
            this.btnAsyncWrite.Text = "Async Write";
            this.btnAsyncWrite.UseVisualStyleBackColor = false;
            this.btnAsyncWrite.Click += new System.EventHandler(this.btnAsyncWrite_Click);
            // 
            // btnAsyncRead
            // 
            this.btnAsyncRead.BackColor = System.Drawing.SystemColors.Control;
            this.btnAsyncRead.Cursor = System.Windows.Forms.Cursors.Default;
            this.btnAsyncRead.Enabled = false;
            this.btnAsyncRead.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnAsyncRead.ForeColor = System.Drawing.SystemColors.ControlText;
            this.btnAsyncRead.Location = new System.Drawing.Point(235, 190);
            this.btnAsyncRead.Name = "btnAsyncRead";
            this.btnAsyncRead.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.btnAsyncRead.Size = new System.Drawing.Size(125, 37);
            this.btnAsyncRead.TabIndex = 101;
            this.btnAsyncRead.Text = "Async Read";
            this.btnAsyncRead.UseVisualStyleBackColor = false;
            this.btnAsyncRead.Click += new System.EventHandler(this.btnAsyncRead_Click);
            // 
            // btnMaxAge
            // 
            this.btnMaxAge.BackColor = System.Drawing.SystemColors.Control;
            this.btnMaxAge.Cursor = System.Windows.Forms.Cursors.Default;
            this.btnMaxAge.Enabled = false;
            this.btnMaxAge.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnMaxAge.ForeColor = System.Drawing.SystemColors.ControlText;
            this.btnMaxAge.Location = new System.Drawing.Point(97, 146);
            this.btnMaxAge.Name = "btnMaxAge";
            this.btnMaxAge.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.btnMaxAge.Size = new System.Drawing.Size(125, 37);
            this.btnMaxAge.TabIndex = 97;
            this.btnMaxAge.Text = "MaxAge ON";
            this.btnMaxAge.UseVisualStyleBackColor = false;
            this.btnMaxAge.Click += new System.EventHandler(this.btnMaxAge_Click);
            // 
            // btnAdvise
            // 
            this.btnAdvise.BackColor = System.Drawing.SystemColors.Control;
            this.btnAdvise.Cursor = System.Windows.Forms.Cursors.Default;
            this.btnAdvise.Enabled = false;
            this.btnAdvise.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnAdvise.ForeColor = System.Drawing.SystemColors.ControlText;
            this.btnAdvise.Location = new System.Drawing.Point(97, 190);
            this.btnAdvise.Name = "btnAdvise";
            this.btnAdvise.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.btnAdvise.Size = new System.Drawing.Size(125, 37);
            this.btnAdvise.TabIndex = 100;
            this.btnAdvise.Text = "Advise";
            this.btnAdvise.UseVisualStyleBackColor = false;
            this.btnAdvise.Click += new System.EventHandler(this.btnAdvise_Click);
            // 
            // btnWrite
            // 
            this.btnWrite.BackColor = System.Drawing.SystemColors.Control;
            this.btnWrite.Cursor = System.Windows.Forms.Cursors.Default;
            this.btnWrite.Enabled = false;
            this.btnWrite.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnWrite.ForeColor = System.Drawing.SystemColors.ControlText;
            this.btnWrite.Location = new System.Drawing.Point(368, 146);
            this.btnWrite.Name = "btnWrite";
            this.btnWrite.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.btnWrite.Size = new System.Drawing.Size(125, 37);
            this.btnWrite.TabIndex = 99;
            this.btnWrite.Text = "Write";
            this.btnWrite.UseVisualStyleBackColor = false;
            this.btnWrite.Click += new System.EventHandler(this.btnWrite_Click);
            // 
            // btnRead
            // 
            this.btnRead.BackColor = System.Drawing.SystemColors.Control;
            this.btnRead.Cursor = System.Windows.Forms.Cursors.Default;
            this.btnRead.Enabled = false;
            this.btnRead.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnRead.ForeColor = System.Drawing.SystemColors.ControlText;
            this.btnRead.Location = new System.Drawing.Point(235, 146);
            this.btnRead.Name = "btnRead";
            this.btnRead.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.btnRead.Size = new System.Drawing.Size(125, 37);
            this.btnRead.TabIndex = 98;
            this.btnRead.Text = "Read";
            this.btnRead.UseVisualStyleBackColor = false;
            this.btnRead.Click += new System.EventHandler(this.btnRead_Click);
            // 
            // btnConnect
            // 
            this.btnConnect.BackColor = System.Drawing.SystemColors.Control;
            this.btnConnect.Cursor = System.Windows.Forms.Cursors.Default;
            this.btnConnect.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnConnect.ForeColor = System.Drawing.SystemColors.ControlText;
            this.btnConnect.Location = new System.Drawing.Point(368, 101);
            this.btnConnect.Name = "btnConnect";
            this.btnConnect.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.btnConnect.Size = new System.Drawing.Size(125, 38);
            this.btnConnect.TabIndex = 96;
            this.btnConnect.Text = "Connect";
            this.btnConnect.UseVisualStyleBackColor = false;
            this.btnConnect.Click += new System.EventHandler(this.btnConnect_Click);
            // 
            // Label9
            // 
            this.Label9.BackColor = System.Drawing.SystemColors.Control;
            this.Label9.Cursor = System.Windows.Forms.Cursors.Default;
            this.Label9.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Label9.ForeColor = System.Drawing.SystemColors.ControlText;
            this.Label9.Location = new System.Drawing.Point(17, 75);
            this.Label9.Name = "Label9";
            this.Label9.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.Label9.Size = new System.Drawing.Size(96, 18);
            this.Label9.TabIndex = 115;
            this.Label9.Text = "Group Name";
            this.Label9.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // Label8
            // 
            this.Label8.BackColor = System.Drawing.SystemColors.Control;
            this.Label8.Cursor = System.Windows.Forms.Cursors.Default;
            this.Label8.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Label8.ForeColor = System.Drawing.SystemColors.ControlText;
            this.Label8.Location = new System.Drawing.Point(17, 16);
            this.Label8.Name = "Label8";
            this.Label8.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.Label8.Size = new System.Drawing.Size(66, 19);
            this.Label8.TabIndex = 110;
            this.Label8.Text = "Node";
            this.Label8.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // Label6
            // 
            this.Label6.BackColor = System.Drawing.SystemColors.Control;
            this.Label6.Cursor = System.Windows.Forms.Cursors.Default;
            this.Label6.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Label6.ForeColor = System.Drawing.SystemColors.ControlText;
            this.Label6.Location = new System.Drawing.Point(17, 103);
            this.Label6.Name = "Label6";
            this.Label6.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.Label6.Size = new System.Drawing.Size(89, 19);
            this.Label6.TabIndex = 108;
            this.Label6.Text = "Update Rate";
            this.Label6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // Label5
            // 
            this.Label5.BackColor = System.Drawing.SystemColors.Control;
            this.Label5.Cursor = System.Windows.Forms.Cursors.Default;
            this.Label5.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Label5.ForeColor = System.Drawing.SystemColors.ControlText;
            this.Label5.Location = new System.Drawing.Point(17, 44);
            this.Label5.Name = "Label5";
            this.Label5.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.Label5.Size = new System.Drawing.Size(96, 19);
            this.Label5.TabIndex = 107;
            this.Label5.Text = "Server Name";
            this.Label5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // Label4
            // 
            this.Label4.BackColor = System.Drawing.SystemColors.Control;
            this.Label4.Cursor = System.Windows.Forms.Cursors.Default;
            this.Label4.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Label4.ForeColor = System.Drawing.SystemColors.ControlText;
            this.Label4.Location = new System.Drawing.Point(445, 240);
            this.Label4.Name = "Label4";
            this.Label4.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.Label4.Size = new System.Drawing.Size(58, 19);
            this.Label4.TabIndex = 106;
            this.Label4.Text = "Quality";
            // 
            // Label3
            // 
            this.Label3.BackColor = System.Drawing.SystemColors.Control;
            this.Label3.Cursor = System.Windows.Forms.Cursors.Default;
            this.Label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Label3.ForeColor = System.Drawing.SystemColors.ControlText;
            this.Label3.Location = new System.Drawing.Point(328, 240);
            this.Label3.Name = "Label3";
            this.Label3.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.Label3.Size = new System.Drawing.Size(78, 19);
            this.Label3.TabIndex = 105;
            this.Label3.Text = "Date/Time";
            // 
            // Label2
            // 
            this.Label2.BackColor = System.Drawing.SystemColors.Control;
            this.Label2.Cursor = System.Windows.Forms.Cursors.Default;
            this.Label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Label2.ForeColor = System.Drawing.SystemColors.ControlText;
            this.Label2.Location = new System.Drawing.Point(198, 240);
            this.Label2.Name = "Label2";
            this.Label2.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.Label2.Size = new System.Drawing.Size(49, 19);
            this.Label2.TabIndex = 104;
            this.Label2.Text = "Value";
            // 
            // Label1
            // 
            this.Label1.BackColor = System.Drawing.SystemColors.Control;
            this.Label1.Cursor = System.Windows.Forms.Cursors.Default;
            this.Label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Label1.ForeColor = System.Drawing.SystemColors.ControlText;
            this.Label1.Location = new System.Drawing.Point(36, 240);
            this.Label1.Name = "Label1";
            this.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.Label1.Size = new System.Drawing.Size(78, 19);
            this.Label1.TabIndex = 103;
            this.Label1.Text = "Item Name";
            // 
            // txtNode
            // 
            this.txtNode.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtNode.Location = new System.Drawing.Point(114, 14);
            this.txtNode.Name = "txtNode";
            this.txtNode.Size = new System.Drawing.Size(160, 20);
            this.txtNode.TabIndex = 111;
            this.txtNode.Text = "localhost";
            // 
            // cmbSvrName
            // 
            this.cmbSvrName.FormattingEnabled = true;
            this.cmbSvrName.Location = new System.Drawing.Point(114, 44);
            this.cmbSvrName.Name = "cmbSvrName";
            this.cmbSvrName.Size = new System.Drawing.Size(160, 20);
            this.cmbSvrName.TabIndex = 112;
            this.cmbSvrName.Text = "Takebishi.Dxp.1";
            // 
            // btnRefresh
            // 
            this.btnRefresh.BackColor = System.Drawing.SystemColors.Control;
            this.btnRefresh.Cursor = System.Windows.Forms.Cursors.Default;
            this.btnRefresh.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnRefresh.ForeColor = System.Drawing.SystemColors.ControlText;
            this.btnRefresh.Location = new System.Drawing.Point(286, 14);
            this.btnRefresh.Name = "btnRefresh";
            this.btnRefresh.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.btnRefresh.Size = new System.Drawing.Size(74, 23);
            this.btnRefresh.TabIndex = 113;
            this.btnRefresh.Text = "Refresh";
            this.btnRefresh.UseVisualStyleBackColor = false;
            this.btnRefresh.Click += new System.EventHandler(this.btnRefresh_Click);
            // 
            // txtGrp
            // 
            this.txtGrp.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtGrp.Location = new System.Drawing.Point(114, 72);
            this.txtGrp.Name = "txtGrp";
            this.txtGrp.Size = new System.Drawing.Size(108, 20);
            this.txtGrp.TabIndex = 114;
            this.txtGrp.Text = "Group1";
            this.txtGrp.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Location = new System.Drawing.Point(326, 649);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(164, 12);
            this.label10.TabIndex = 116;
            this.label10.Text = "VCDotNetRcwSample ver2.3";
            // 
            // chkGrpActive
            // 
            this.chkGrpActive.AutoSize = true;
            this.chkGrpActive.Checked = true;
            this.chkGrpActive.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkGrpActive.Enabled = false;
            this.chkGrpActive.Location = new System.Drawing.Point(396, 49);
            this.chkGrpActive.Name = "chkGrpActive";
            this.chkGrpActive.Size = new System.Drawing.Size(92, 16);
            this.chkGrpActive.TabIndex = 117;
            this.chkGrpActive.Text = "GroupActive";
            this.chkGrpActive.UseVisualStyleBackColor = true;
            this.chkGrpActive.CheckedChanged += new System.EventHandler(this.chkGrpActive_CheckedChanged);
            // 
            // rdoDA30
            // 
            this.rdoDA30.Checked = true;
            this.rdoDA30.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.rdoDA30.Location = new System.Drawing.Point(436, 75);
            this.rdoDA30.Name = "rdoDA30";
            this.rdoDA30.Size = new System.Drawing.Size(57, 18);
            this.rdoDA30.TabIndex = 95;
            this.rdoDA30.TabStop = true;
            this.rdoDA30.Text = "3.0";
            this.rdoDA30.CheckedChanged += new System.EventHandler(this.rdoDA30_CheckedChanged);
            // 
            // rdoDA20
            // 
            this.rdoDA20.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.rdoDA20.Location = new System.Drawing.Point(368, 75);
            this.rdoDA20.Name = "rdoDA20";
            this.rdoDA20.Size = new System.Drawing.Size(58, 18);
            this.rdoDA20.TabIndex = 94;
            this.rdoDA20.Text = "2.0";
            this.rdoDA20.CheckedChanged += new System.EventHandler(this.rdoDA20_CheckedChanged);
            // 
            // Label7
            // 
            this.Label7.BackColor = System.Drawing.SystemColors.Control;
            this.Label7.Cursor = System.Windows.Forms.Cursors.Default;
            this.Label7.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Label7.ForeColor = System.Drawing.SystemColors.ControlText;
            this.Label7.Location = new System.Drawing.Point(253, 75);
            this.Label7.Name = "Label7";
            this.Label7.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.Label7.Size = new System.Drawing.Size(108, 18);
            this.Label7.TabIndex = 109;
            this.Label7.Text = "OPCDA Version";
            this.Label7.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // Form1
            // 
            this.AutoScaleBaseSize = new System.Drawing.Size(6, 14);
            this.ClientSize = new System.Drawing.Size(525, 575);
            this.Controls.Add(this.chkGrpActive);
            this.Controls.Add(this.label10);
            this.Controls.Add(this.Label9);
            this.Controls.Add(this.txtGrp);
            this.Controls.Add(this.btnRefresh);
            this.Controls.Add(this.cmbSvrName);
            this.Controls.Add(this.txtNode);
            this.Controls.Add(this.Label8);
            this.Controls.Add(this.txtUpdateRate);
            this.Controls.Add(this.rdoDA30);
            this.Controls.Add(this.rdoDA20);
            this.Controls.Add(this.Label7);
            this.Controls.Add(this.Label6);
            this.Controls.Add(this.btnAsyncWrite);
            this.Controls.Add(this.btnAsyncRead);
            this.Controls.Add(this.btnMaxAge);
            this.Controls.Add(this.btnAdvise);
            this.Controls.Add(this.btnWrite);
            this.Controls.Add(this.btnRead);
            this.Controls.Add(this.btnConnect);
            this.Controls.Add(this.Label5);
            this.Controls.Add(this.Label4);
            this.Controls.Add(this.Label3);
            this.Controls.Add(this.Label2);
            this.Controls.Add(this.Label1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.Name = "Form1";
            this.Text = "Form1";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

		}
		#endregion

		/// <summary>
		/// Main entry point of application
		/// </summary>
		[STAThread]
		static void Main() 
		{
			Application.Run(new Form1());
		}


		private void Copy(DEF_COPY_DIR Direction)
		{
			int i, j;
			string sBufGet;
			string[] sBufPut;
			switch (Direction) 
			{
				case DEF_COPY_DIR.DISP_TO_MEM:
					for (i = 0; i < ITEMMAX; i++) 
					{
						if (oVal[i] is Array) 
						{
							//erase sBufPut
							sBufPut = txtValue[i].Text.Split(',');
							// get temporal array pointer
							Array ary = (Array)oVal[i];
							object[] oAry = new object[ary.Length];
							for (j = 0; j < ary.Length; j++)
							{
								if (j >= sBufPut.Length)
									break;

								// confirm change double (as "IsNumeric")
								//double d;
								//if (double.TryParse(sBufPut[j],
								//	System.Globalization.NumberStyles.Any,
								//	System.Globalization.NumberFormatInfo.InvariantInfo,
								//	out d))
								try
								{
									if (oVal[i] is UInt16)
									{
										oAry[j] = UInt16.Parse(sBufPut[j]);
									}
									else if (oVal[i] is UInt32)
									{
										oAry[j] = UInt32.Parse(sBufPut[j]);
									}
									else 
									{
										oAry[j] = (object)sBufPut[j];
									}
								}
								catch
								{
								}
							}
							oVal[i] = oAry;
						}
						else
						{
							// confirm change double (as "IsNumeric")
							try
							{
								if (oVal[i] is UInt16)
								{
									oVal[i] = UInt16.Parse(txtValue[i].Text);
								}
								else if (oVal[i] is UInt32)
								{
									oVal[i] = UInt32.Parse(txtValue[i].Text);
								}
								else 
								{
									oVal[i] = txtValue[i].Text;
								}
							}
							catch
							{
							}
						}
						// confirm change DateTime (as "IsDate")
						try
						{
							DateTime dt = DateTime.Parse(txtTime[i].Text);
							JointInt joFt = new JointInt();
							joFt.lJoint = dt.ToFileTime();
							dTime[i].dwHighDateTime = joFt.iHi;
							dTime[i].dwLowDateTime = joFt.iLo;
						}
						catch
						{
						}
						// confirm change double (as "IsNumeric")
						try
						{
							wQuality[i] = short.Parse(txtQuality[i].Text);
						}
						catch
						{
						}
					}
					break;

				case DEF_COPY_DIR.MEM_TO_DISP:
					for (i = 0; i < ITEMMAX; i++) 
					{
						if (oVal[i] != null) 
						{		// 05/03/10 If object is nothing, do not read value.
							if (oVal[i] is Array)
							{
								sBufGet = "";
								Array ary = (Array)oVal[i];
								for (j = 0; j < ary.Length; j++)
								{
									sBufGet = sBufGet + ary.GetValue(j) + ",";	//05/03/10 Change value to string
								}
								txtValue[i].Text = sBufGet;
							}
							else 
							{
								txtValue[i].Text = oVal[i].ToString();
							}
						}
//						if (dTime[i] != null)
							JointInt joFt = new JointInt();
							joFt.iHi = dTime[i].dwHighDateTime;
							joFt.iLo = dTime[i].dwLowDateTime;
							DateTime dt = DateTime.FromFileTime(joFt.lJoint);
							txtTime[i].Text = dt.ToString();
//							long lftH = dTime[i].dwHighDateTime;
//							long lftL = (dTime[i].dwLowDateTime >= 0) ? dTime[i].dwLowDateTime : dTime[i].dwLowDateTime + (long)0x100000000;
//							long ljoint = (lftH << 32) + lftL;
//							DateTime dt = DateTime.FromFileTime(ljoint);
//							txtTime[i].Text = dt.ToString();
//						if (wQuality[i] != 0)
							txtQuality[i].Text = wQuality[i].ToString();
					}
					break;
			}
		}

		private void Copy(int iItemCount,
						int[] iClientHds,
						object[] vValues,
						OpcRcw.Da.FILETIME[] ftTimeStamps,
						short[] wQualities,
						int[] pErrors)
		{
			int i, j;
			for (i = 0; i < iItemCount; i++)
			{
				for (j = 0; j < ITEMMAX; j++) 
				{
					if (cH[j] == iClientHds[i])
					{
						if (pErrors[i] == 0)
						{
							oVal[j] = vValues[i];
							dTime[j] = ftTimeStamps[i];
							wQuality[j] = wQualities[i];
						}
						break;
					}
				}
			}
		}

		private void Form1_Load(object sender, System.EventArgs e)
		{
			// create TextBox
			for (int i = 0; i < ITEMMAX; i++)
			{
				txtItemName[i] = new System.Windows.Forms.TextBox();
				txtItemName[i].AutoSize = false;
				txtItemName[i].Top = VAL_CTRL_TOP + (VAL_ITEMNAME_HEIGHT + VAL_CTRL_SPACE) * (i+1);
				txtItemName[i].Left = VAL_CTRL_LEFT;
				txtItemName[i].Height = VAL_ITEMNAME_HEIGHT;
				txtItemName[i].Width = VAL_ITEMNAME_WIDTH;
				this.Controls.Add(txtItemName[i]);

				txtValue[i] = new System.Windows.Forms.TextBox();
				txtValue[i].AutoSize = false;
                txtValue[i].Top = VAL_CTRL_TOP + (VAL_VALUE_HEIGHT + VAL_CTRL_SPACE) * (i + 1);
				txtValue[i].Left = VAL_CTRL_LEFT + (VAL_ITEMNAME_WIDTH + VAL_CTRL_SPACE);
				txtValue[i].Height = VAL_VALUE_HEIGHT;
				txtValue[i].Width = VAL_VALUE_WIDTH;
				this.Controls.Add(txtValue[i]);

				txtTime[i] = new System.Windows.Forms.TextBox();
				txtTime[i].AutoSize = false;
                txtTime[i].Top = VAL_CTRL_TOP + (VAL_TIME_HEIGHT + VAL_CTRL_SPACE) * (i + 1);
				txtTime[i].Left = VAL_CTRL_LEFT + (VAL_ITEMNAME_WIDTH + VAL_CTRL_SPACE) + (VAL_VALUE_WIDTH + VAL_CTRL_SPACE);
				txtTime[i].Height = VAL_TIME_HEIGHT;
				txtTime[i].Width = VAL_TIME_WIDTH;
				this.Controls.Add(txtTime[i]);

				txtQuality[i] = new System.Windows.Forms.TextBox();
				txtQuality[i].AutoSize = false;
                txtQuality[i].Top = VAL_CTRL_TOP + (VAL_QUALITY_HEIGHT + VAL_CTRL_SPACE) * (i + 1);
				txtQuality[i].Left = VAL_CTRL_LEFT + (VAL_ITEMNAME_WIDTH + VAL_CTRL_SPACE) + (VAL_VALUE_WIDTH + VAL_CTRL_SPACE) + (VAL_TIME_WIDTH + VAL_CTRL_SPACE);
				txtQuality[i].Height = VAL_QUALITY_HEIGHT;
				txtQuality[i].Width = VAL_QUALITY_WIDTH;
				this.Controls.Add(txtQuality[i]);
			}

			// Create an instance of StreamReader to read from a file.
			// The using statement also closes the StreamReader.
			spFolder = System.Environment.GetFolderPath(Environment.SpecialFolder.Personal) +"\\SEFA\\OPC\\Link";
			csInifilePath = spFolder + "\\" + OLGA_OPC_INI;

			Boolean bErr = false;
			try
			{
				if (System.IO.File.Exists(csInifilePath))
				{
					StreamReader sr = new StreamReader(csInifilePath);
					String line;
					// Read and display lines from the file until the end of 
					// the file is reached.
					txtNode.Text = sr.ReadLine();
					cmbSvrName.Text = sr.ReadLine();
					txtGrp.Text = sr.ReadLine();
					txtUpdateRate.Text = sr.ReadLine();
					OpcdaVer = (DEF_OPCDA)int.Parse(sr.ReadLine());
					if (OpcdaVer == DEF_OPCDA.VER_30)
					{
						rdoDA30.Checked = true;
					}
					else
					{
						rdoDA20.Checked = true;
					}

					for (int i = 0; i < ITEMMAX; i++)
					{
						txtItemName[i].Text = sr.ReadLine();
					}

					while ((line = sr.ReadLine()) != null)
					{
						Console.WriteLine(line);
					}
					sr.Close();
					btnRefresh_Click(null, null);
				}
				else
				{
					bErr = true;
				}
			}
			catch (Exception exc)
			{
				bErr = true;
				Console.WriteLine(exc.ToString());
			}

			if ( bErr == true )
			{
				txtNode.Text = "localhost";
                cmbSvrName.Text = "SPT.OLGAOPCServer.7";
				txtGrp.Text = "Sim.SNU_OLGA";
				txtUpdateRate.Text = "100000";
				OpcdaVer = DEF_OPCDA.VER_30;
				rdoDA30.Checked = true;
				/*for (int i = 0; i < ITEMMAX; i++)
				{
					txtItemName[i].Text = System.String.Format("Device1.D{0}", i);
				}
                */
                txtItemName[0].Text = "Sim.SNU_OLGA.ExternalClock";
                txtItemName[1].Text = "Sim.SNU_OLGA.TIME";
                txtItemName[2].Text = "Sim.SNU_OLGA.CPF 1.PRESSURE";
                txtItemName[3].Text = "Sim.SNU_OLGA.CPF 1.PTBOU";
                txtItemName[4].Text = "Sim.SNU_OLGA.CPF 1.TMBOU";
                txtItemName[5].Text = "Sim.SNU_OLGA.CPF 1.GTBOU";
                txtItemName[6].Text = "Sim.SNU_OLGA.CPF 1.CGGBOU";
                txtItemName[7].Text = "Sim.SNU_OLGA.CPF 1.CGLTHLBOU";
                txtItemName[8].Text = "Sim.SNU_OLGA.CPF 1.CGLTWTBOU";
			}
		}

		protected override void OnClosed(EventArgs e)
		{
			int i;

			if (OPCSvr != null) 
			{
				if (OPCSvr.IsConnect() == true) 
				{
					btnConnect_Click(btnConnect, new System.EventArgs());
				}

				try 
				{
					if (!System.IO.File.Exists(csInifilePath))
					{
						System.IO.Directory.CreateDirectory(spFolder);
						using (System.IO.FileStream hStream = System.IO.File.Create(csInifilePath))
						{
							// files tream close
							if (hStream != null)
							{
								hStream.Close();
							}
						}
					}
					using (StreamWriter sw = new StreamWriter(csInifilePath)) // when closing, save the state in the given file.
					{
						// Add some text to the file.
						sw.WriteLine(txtNode.Text);
						sw.WriteLine(cmbSvrName.Text);
						sw.WriteLine(txtGrp.Text);
						sw.WriteLine(txtUpdateRate.Text);
						sw.WriteLine((int)OpcdaVer);
						for (i = 0; i < ITEMMAX; i++) 
						{
							sw.WriteLine(txtItemName[i].Text);
						}
						sw.Close();
					}
				}
				catch (Exception exc) 
				{
					Console.WriteLine(exc.ToString());
				}

				OPCSvr = null;
			}

			base.OnClosed (e);
		}

		private void btnRefresh_Click(object sender, EventArgs e)
		{
			List<string> listServerNameList = new List<string>();
			OPCSvr.BrowseServerNameList(OpcdaVer, txtNode.Text, ref listServerNameList);

			// Add Server Name List
			cmbSvrName.Items.Clear();
			for (int n = 0; n < listServerNameList.Count; n++)
				cmbSvrName.Items.Add(listServerNameList[n]);

			listServerNameList.Clear();
			return;
		}

		private void btnConnect_Click(object sender, System.EventArgs e)
		{
			int i;
			if (cmbSvrName.Text == "")
			{
				MessageBox.Show("Server Name is not registed.", "btnConnect_Click");
				return;
			}

			// confirm change double (as "IsNumeric")
			double d;
			if (!double.TryParse(txtUpdateRate.Text,
					System.Globalization.NumberStyles.Any,
					System.Globalization.NumberFormatInfo.InvariantInfo,
					out d))
			{
				MessageBox.Show("Update Rate is unsuitable.", "btnConnect_Click");
				return;
			}

			if (OpcdaVer == DEF_OPCDA.VER_NONE) {
				MessageBox.Show("Select OPCDA Version.", "btnConnect_Click");
				return;
			}

			if (OPCSvr.IsConnect() == false) 
			{
				// --- connect OPCServer
				if (OPCSvr.Connect(OpcdaVer, txtNode.Text, cmbSvrName.Text, txtGrp.Text, int.Parse(txtUpdateRate.Text)) == true) //if connection is done but items are not added
				{
					for (i = 0; i < ITEMMAX; i++) 
					{
						sItemName[i] = txtItemName[i].Text;
						cH[i] = i;
					}
					if (OPCSvr.AddItem(sItemName, cH, sH) == false) 
					{
						MessageBox.Show("Registing Items is failed.", "btnConnect_Click");
						OPCSvr.Disconnect();
						return;
					}
				}
				else // if connection is not done
				{
					MessageBox.Show("Connecting is failed.", "btnConnect_Click");
					OPCSvr.Disconnect();
					return;
				}
				cmbSvrName.Enabled = false;
				txtUpdateRate.Enabled = false;
				rdoDA20.Enabled = false;
				rdoDA30.Enabled = false;
				cmbSvrName.Enabled = false;
				txtNode.Enabled = false;
				txtGrp.Enabled = false;
				btnRefresh.Enabled = false;
				btnConnect.Text = "Disconnect";
				btnAdvise.Enabled = true;
				btnAdvise.Text = "Advise";
				if (OpcdaVer == DEF_OPCDA.VER_30)
				{
					btnMaxAge.Enabled = true;
					btnMaxAge.Text = "MaxAge ON";
				}
				else 
				{
					btnMaxAge.Enabled = false;
				}
				btnRead.Enabled = true;
				btnWrite.Enabled = true;
				btnAsyncRead.Enabled = false;
				btnAsyncWrite.Enabled = false;
                chkGrpActive.Enabled = true;
				for (i = 0; i < ITEMMAX; i++)
				{
					txtItemName[i].Enabled = false;
				}
				OPCSvr.SyncRead(OPCDATASOURCE.OPC_DS_DEVICE, sH, oVal, dTime, wQuality);
				Copy(DEF_COPY_DIR.MEM_TO_DISP);
			}
			else 
			{
				// --- disconnect OPCServer
				OPCSvr.Disconnect();
				cmbSvrName.Enabled = true;
				txtUpdateRate.Enabled = true;
				rdoDA20.Enabled = true;
				rdoDA30.Enabled = true;
				cmbSvrName.Enabled = true;
				txtGrp.Enabled = true;
				btnRefresh.Enabled = true;
				txtNode.Enabled = true;
				btnConnect.Text = "Connect";
				btnAdvise.Enabled = false;
				btnMaxAge.Enabled = false;
				btnRead.Enabled = false;
				btnWrite.Enabled = false;
				btnAsyncRead.Enabled = false;
				btnAsyncWrite.Enabled = false;
                chkGrpActive.Enabled = false;
				for (i = 0; i < ITEMMAX; i++)
				{
					txtItemName[i].Enabled = true;
				}
//				tmMaxAge.Enabled = false;
			}
		}

		private void btnMaxAge_Click(object sender, System.EventArgs e)
		{
			OPCSvr.ReadMaxAge(10000, sH, oVal, dTime, wQuality);
			//OPCSvr.ReadMaxAge(CInt(txtUpdateRate.Text), sH, oVal, dTime, wQuality);
			Copy(DEF_COPY_DIR.MEM_TO_DISP);
		}

		private void btnRead_Click(object sender, System.EventArgs e)
		{
			OPCSvr.SyncRead(OPCDATASOURCE.OPC_DS_DEVICE, sH, oVal, dTime, wQuality);
			Copy(DEF_COPY_DIR.MEM_TO_DISP);
		}

		private void btnWrite_Click(object sender, System.EventArgs e)
		{
			Copy(DEF_COPY_DIR.DISP_TO_MEM);
			OPCSvr.SyncWrite(sH, oVal);
		}

		private void btnAdvise_Click(object sender, System.EventArgs e)
		{
			if (OPCSvr.IsAdvise() == false)
			{
				OPCSvr.Advise();
				btnAdvise.Text = "Unadvise";
				btnAsyncRead.Enabled = true;
				btnAsyncWrite.Enabled = true;
			}
			else 
			{
				OPCSvr.Unadvise();
				btnAdvise.Text = "Advise";
				btnAsyncRead.Enabled = false;
				btnAsyncWrite.Enabled = false;
			}
		}

		private void btnAsyncRead_Click(object sender, System.EventArgs e)
		{
			int wTransID = 10000;
			int wCancelID;
			OPCSvr.AsyncRead(wTransID, out wCancelID, sH);
		}

		private void btnAsyncWrite_Click(object sender, System.EventArgs e)
		{
			int wTransID = 20000;
			int wCancelID;
			Copy(DEF_COPY_DIR.DISP_TO_MEM);
			OPCSvr.AsyncWrite(wTransID, out wCancelID, sH, oVal);
		}

		private void rdoDA20_CheckedChanged(object sender, System.EventArgs e)
		{
			OpcdaVer = DEF_OPCDA.VER_20;
		}

		private void rdoDA30_CheckedChanged(object sender, System.EventArgs e)
		{
			OpcdaVer = DEF_OPCDA.VER_30;
		}

		private void OPCSvr_DataChange(
			int						wTransID,
			int						iItemCount,
			int[]					iClientHds,
			object[]				vValues,
			OpcRcw.Da.FILETIME[]	ftTimeStamps,
			short[]					wQualities,
			int[]					pErrors)
		{
			Copy(iItemCount, iClientHds, vValues, ftTimeStamps, wQualities, pErrors);
			this.threadCopy = new Thread(new ThreadStart(this.ThreadProcCopy));   // 06/09/20 for VS2005
			this.threadCopy.Start();
			//Copy(DEF_COPY_DIR.MEM_TO_DISP);
		}

		private void OPCSvr_ReadComplete(
			int						wTransID,
			int						iItemCount,
			int[]					iClientHds,
			object[]				vValues,
			OpcRcw.Da.FILETIME[]	ftTimeStamps,
			short[]					wQualities,
			int[]					pErrors)
		{
			Copy(iItemCount, iClientHds, vValues, ftTimeStamps, wQualities, pErrors);
			this.threadCopy = new Thread(new ThreadStart(this.ThreadProcCopy));   // 06/09/20 for VS2005
			this.threadCopy.Start();
			//Copy(DEF_COPY_DIR.MEM_TO_DISP);
		}

		private void OPCSvr_WriteComplete(
			int						wTransID,
			int						iItemCount,
			int[]					iClientHds,
			int[]					pErrors)
		{
		}

		private void OPCSvr_CancelComplete(
			int						wTransID)
		{
		}

        // 2011/11/14 シャット?ウンイベントを受けれるようにする	(
        //===================================================================
        //	Function		: OPCSvr_ShutDown(CString sReason)
        //	Distribution	: シャット?ウンイベント
        //===================================================================
        private void OPCSvr_ShutDown(String sReason)
        {
            MessageBox.Show("Receive ShutDown Event [Reason = " + sReason + "]\r\nPlease Disconnect!");
        }
        // 2011/11/14 シャット?ウンイベントを受けれるようにする	)


		/************************/
		private void ThreadProcCopy() // 06/09/20 for VS2005
		{
			for (int i = 0; i < ITEMMAX; i++)
			{
				this.SetValue(i);
				this.SetQuality(i);
				this.SetTime(i);
			}
		}
		private void SetValue(int i)        // 06/09/20 for VS2005
		{
			// InvokeRequired required compares the thread ID of the
			// calling thread to the thread ID of the creating thread.
			// If these threads are different, it returns true.
			if (this.txtValue[i].InvokeRequired)
			{
				SetValueCallback d = new SetValueCallback(SetValue);
				this.Invoke(d, new object[] { i });
			}
			else
			{
				string sBufGet;

				if (oVal[i] != null)
				{		// 05/03/10 If object is nothing, do not read value.
					if (oVal[i] is Array)
					{
						sBufGet = "";
						Array ary = (Array)oVal[i];
						for (int j = 0; j < ary.Length; j++)
						{
							sBufGet = sBufGet + ary.GetValue(j) + ",";	//05/03/10 Change value to string
						}
						txtValue[i].Text = sBufGet;
					}
					else
					{
						txtValue[i].Text = oVal[i].ToString();
					}
				}
			}
		}
		private void SetTime(int i)   // 06/09/20 for VS2005
		{
			// InvokeRequired required compares the thread ID of the
			// calling thread to the thread ID of the creating thread.
			// If these threads are different, it returns true.
			if (this.txtTime[i].InvokeRequired)
			{
				SetTimeCallback d = new SetTimeCallback(SetTime);
				this.Invoke(d, new object[] { i });
			}
			else
			{
				JointInt joFt = new JointInt();
				joFt.iHi = dTime[i].dwHighDateTime;
				joFt.iLo = dTime[i].dwLowDateTime;
				DateTime dt = DateTime.FromFileTime(joFt.lJoint);
				txtTime[i].Text = dt.ToString();
			}
		}
		private void SetQuality(int i)   // 06/09/20 for VS2005
		{
			// InvokeRequired required compares the thread ID of the
			// calling thread to the thread ID of the creating thread.
			// If these threads are different, it returns true.
			if (this.txtQuality[i].InvokeRequired)
			{
				SetQualityCallback d = new SetQualityCallback(SetQuality);
				this.Invoke(d, new object[] { i });
			}
			else
			{
				txtQuality[i].Text = wQuality[i].ToString();
			}
		}

		// 2012/01/12 グル?プのアクティブ/非アクティブを切り替え	(
		private void chkGrpActive_CheckedChanged(object sender, EventArgs e)
		{
			String sErrMsg;
			if (chkGrpActive.Checked)
			{
				// グル?プのアクティブ化
				OPCSvr.SetGroupActiveStatus( true, out sErrMsg);
			}
			else
			{
				// グル?プの非アクティブ化
				OPCSvr.SetGroupActiveStatus( false, out sErrMsg);
			}
		}
		// 2012/01/12 グル?プのアクティブ/非アクティブを切り替え	)
	}
}
