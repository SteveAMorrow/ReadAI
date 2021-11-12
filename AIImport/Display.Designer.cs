
namespace Bentley.Plant.App.Pid.AIImport
{
    partial class Display
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

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

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.listView = new System.Windows.Forms.ListView();
            this.columnHeaderId = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.columnHeaderUserLabel = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.columnHeaderiModelClass = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.columnHeaderSpecificClass = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.buttonCreate = new System.Windows.Forms.Button();
            this.splitContainer = new System.Windows.Forms.SplitContainer();
            this.splitContainerProps = new System.Windows.Forms.SplitContainer();
            this.listViewProps = new System.Windows.Forms.ListView();
            this.columnHeaderName = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.columnHeaderValue = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.tabControl = new System.Windows.Forms.TabControl();
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer)).BeginInit();
            this.splitContainer.Panel1.SuspendLayout();
            this.splitContainer.Panel2.SuspendLayout();
            this.splitContainer.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.splitContainerProps)).BeginInit();
            this.splitContainerProps.Panel1.SuspendLayout();
            this.splitContainerProps.Panel2.SuspendLayout();
            this.splitContainerProps.SuspendLayout();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // listView
            // 
            this.listView.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.columnHeaderId,
            this.columnHeaderUserLabel,
            this.columnHeaderiModelClass,
            this.columnHeaderSpecificClass});
            this.listView.Dock = System.Windows.Forms.DockStyle.Fill;
            this.listView.FullRowSelect = true;
            this.listView.HideSelection = false;
            this.listView.Location = new System.Drawing.Point(0, 0);
            this.listView.Name = "listView";
            this.listView.Size = new System.Drawing.Size(771, 282);
            this.listView.TabIndex = 0;
            this.listView.UseCompatibleStateImageBehavior = false;
            this.listView.View = System.Windows.Forms.View.Details;
            this.listView.SelectedIndexChanged += new System.EventHandler(this.listView_SelectedIndexChanged);
            // 
            // columnHeaderId
            // 
            this.columnHeaderId.Text = "Id";
            this.columnHeaderId.Width = 83;
            // 
            // columnHeaderUserLabel
            // 
            this.columnHeaderUserLabel.Text = "User Label";
            this.columnHeaderUserLabel.Width = 146;
            // 
            // columnHeaderiModelClass
            // 
            this.columnHeaderiModelClass.Text = "Parent Class";
            this.columnHeaderiModelClass.Width = 171;
            // 
            // columnHeaderSpecificClass
            // 
            this.columnHeaderSpecificClass.Text = "SpecificClass";
            this.columnHeaderSpecificClass.Width = 143;
            // 
            // buttonCreate
            // 
            this.buttonCreate.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.buttonCreate.Location = new System.Drawing.Point(707, 550);
            this.buttonCreate.Name = "buttonCreate";
            this.buttonCreate.Size = new System.Drawing.Size(63, 28);
            this.buttonCreate.TabIndex = 1;
            this.buttonCreate.Text = "Create";
            this.buttonCreate.UseVisualStyleBackColor = true;
            this.buttonCreate.Click += new System.EventHandler(this.buttonCreate_Click);
            // 
            // splitContainer
            // 
            this.splitContainer.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.splitContainer.Location = new System.Drawing.Point(1, 2);
            this.splitContainer.Name = "splitContainer";
            this.splitContainer.Orientation = System.Windows.Forms.Orientation.Horizontal;
            // 
            // splitContainer.Panel1
            // 
            this.splitContainer.Panel1.Controls.Add(this.listView);
            // 
            // splitContainer.Panel2
            // 
            this.splitContainer.Panel2.Controls.Add(this.splitContainerProps);
            this.splitContainer.Size = new System.Drawing.Size(771, 540);
            this.splitContainer.SplitterDistance = 282;
            this.splitContainer.TabIndex = 3;
            // 
            // splitContainerProps
            // 
            this.splitContainerProps.Dock = System.Windows.Forms.DockStyle.Fill;
            this.splitContainerProps.Location = new System.Drawing.Point(0, 0);
            this.splitContainerProps.Name = "splitContainerProps";
            // 
            // splitContainerProps.Panel1
            // 
            this.splitContainerProps.Panel1.Controls.Add(this.listViewProps);
            // 
            // splitContainerProps.Panel2
            // 
            this.splitContainerProps.Panel2.Controls.Add(this.groupBox1);
            this.splitContainerProps.Size = new System.Drawing.Size(771, 254);
            this.splitContainerProps.SplitterDistance = 301;
            this.splitContainerProps.TabIndex = 0;
            // 
            // listViewProps
            // 
            this.listViewProps.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.columnHeaderName,
            this.columnHeaderValue});
            this.listViewProps.Dock = System.Windows.Forms.DockStyle.Fill;
            this.listViewProps.HideSelection = false;
            this.listViewProps.Location = new System.Drawing.Point(0, 0);
            this.listViewProps.Name = "listViewProps";
            this.listViewProps.Size = new System.Drawing.Size(301, 254);
            this.listViewProps.TabIndex = 0;
            this.listViewProps.UseCompatibleStateImageBehavior = false;
            this.listViewProps.View = System.Windows.Forms.View.Details;
            // 
            // columnHeaderName
            // 
            this.columnHeaderName.Text = "Name";
            this.columnHeaderName.Width = 89;
            // 
            // columnHeaderValue
            // 
            this.columnHeaderValue.Text = "Value";
            this.columnHeaderValue.Width = 95;
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.tabControl);
            this.groupBox1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.groupBox1.Location = new System.Drawing.Point(0, 0);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(466, 254);
            this.groupBox1.TabIndex = 1;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Associated to";
            // 
            // tabControl
            // 
            this.tabControl.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tabControl.Location = new System.Drawing.Point(3, 16);
            this.tabControl.Name = "tabControl";
            this.tabControl.SelectedIndex = 0;
            this.tabControl.Size = new System.Drawing.Size(460, 235);
            this.tabControl.TabIndex = 0;
            // 
            // Display
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(776, 584);
            this.Controls.Add(this.splitContainer);
            this.Controls.Add(this.buttonCreate);
            this.Name = "Display";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Display";
            this.splitContainer.Panel1.ResumeLayout(false);
            this.splitContainer.Panel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer)).EndInit();
            this.splitContainer.ResumeLayout(false);
            this.splitContainerProps.Panel1.ResumeLayout(false);
            this.splitContainerProps.Panel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.splitContainerProps)).EndInit();
            this.splitContainerProps.ResumeLayout(false);
            this.groupBox1.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.ListView listView;
        private System.Windows.Forms.ColumnHeader columnHeaderiModelClass;
        private System.Windows.Forms.ColumnHeader columnHeaderUserLabel;
        private System.Windows.Forms.ColumnHeader columnHeaderSpecificClass;
        private System.Windows.Forms.ColumnHeader columnHeaderId;
        private System.Windows.Forms.Button buttonCreate;
        private System.Windows.Forms.SplitContainer splitContainer;
        private System.Windows.Forms.SplitContainer splitContainerProps;
        private System.Windows.Forms.ListView listViewProps;
        private System.Windows.Forms.ColumnHeader columnHeaderName;
        private System.Windows.Forms.ColumnHeader columnHeaderValue;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.TabControl tabControl;
    }
}