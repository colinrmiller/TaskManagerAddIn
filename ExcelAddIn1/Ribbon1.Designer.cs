﻿using TaskManager;

namespace TaskManagemer
{
    partial class TaskRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public TaskRibbon()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Component Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.btnTask = this.Factory.CreateRibbonButton();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.button1 = this.Factory.CreateRibbonButton();
            this.group3 = this.Factory.CreateRibbonGroup();
            this.button2 = this.Factory.CreateRibbonButton();
            this.btnNewTaskItem = this.Factory.CreateRibbonButton();
            this.btnDeleteRow = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.group2.SuspendLayout();
            this.group3.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.Groups.Add(this.group1);
            this.tab1.Groups.Add(this.group2);
            this.tab1.Groups.Add(this.group3);
            this.tab1.Label = "Hevy";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.btnTask);
            this.group1.Items.Add(this.btnNewTaskItem);
            this.group1.Items.Add(this.btnDeleteRow);
            this.group1.Label = "Data";
            this.group1.Name = "group1";
            // 
            // btnTask
            // 
            this.btnTask.Label = "New Task";
            this.btnTask.Name = "btnTask";
            this.btnTask.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnLoadData_Click);
            // 
            // group2
            // 
            this.group2.Items.Add(this.button1);
            this.group2.Label = "Data";
            this.group2.Name = "group2";
            // 
            // button1
            // 
            this.button1.Label = "Pull Data";
            this.button1.Name = "button1";
            // 
            // group3
            // 
            this.group3.Items.Add(this.button2);
            this.group3.Label = "Data";
            this.group3.Name = "group3";
            // 
            // button2
            // 
            this.button2.Label = "Pull Data";
            this.button2.Name = "button2";
            // 
            // btnNewTaskItem
            // 
            this.btnNewTaskItem.Label = "New Item";
            this.btnNewTaskItem.Name = "btnNewTaskItem";
            // 
            // btnDeleteRow
            // 
            this.btnDeleteRow.Label = "Delete Row";
            this.btnDeleteRow.Name = "btnDeleteRow";
            // 
            // Hevy
            // 
            this.Name = "Hevy";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            this.group3.ResumeLayout(false);
            this.group3.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnTask;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnNewTaskItem;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnDeleteRow;
    }

    partial class ThisRibbonCollection
    {
        internal TaskRibbon Ribbon1
        {
            get { return this.GetRibbon<TaskRibbon>(); }
        }
    }
}
