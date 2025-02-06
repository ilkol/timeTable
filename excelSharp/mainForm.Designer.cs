namespace excelSharp
{
    partial class mainForm
    {
        /// <summary>
        /// Обязательная переменная конструктора.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Освободить все используемые ресурсы.
        /// </summary>
        /// <param name="disposing">истинно, если управляемый ресурс должен быть удален; иначе ложно.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Код, автоматически созданный конструктором форм Windows

        /// <summary>
        /// Требуемый метод для поддержки конструктора — не изменяйте 
        /// содержимое этого метода с помощью редактора кода.
        /// </summary>
        private void InitializeComponent()
        {
            this.button1 = new System.Windows.Forms.Button();
            this.groupsListBox = new System.Windows.Forms.ComboBox();
            this.createTable = new System.Windows.Forms.Button();
            this.groupListBox = new System.Windows.Forms.TextBox();
            this.addGroupButton = new System.Windows.Forms.Button();
            this.removeGroupButton = new System.Windows.Forms.Button();
            this.timeTableButton = new System.Windows.Forms.Button();
            this.timeTableTextBox = new System.Windows.Forms.TextBox();
            this.writeTimetableButton = new System.Windows.Forms.Button();
            this.changeGroupList = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(951, 50);
            this.button1.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(100, 28);
            this.button1.TabIndex = 0;
            this.button1.Text = "button1";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // groupsListBox
            // 
            this.groupsListBox.FormattingEnabled = true;
            this.groupsListBox.Location = new System.Drawing.Point(16, 30);
            this.groupsListBox.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.groupsListBox.Name = "groupsListBox";
            this.groupsListBox.Size = new System.Drawing.Size(259, 24);
            this.groupsListBox.TabIndex = 1;
            this.groupsListBox.SelectedIndexChanged += new System.EventHandler(this.groupsListBox_SelectedIndexChanged);
            // 
            // createTable
            // 
            this.createTable.Location = new System.Drawing.Point(887, 86);
            this.createTable.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.createTable.Name = "createTable";
            this.createTable.Size = new System.Drawing.Size(164, 28);
            this.createTable.TabIndex = 2;
            this.createTable.Text = "Создать файл";
            this.createTable.UseVisualStyleBackColor = true;
            this.createTable.Click += new System.EventHandler(this.createTable_Click);
            // 
            // groupListBox
            // 
            this.groupListBox.Location = new System.Drawing.Point(16, 63);
            this.groupListBox.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.groupListBox.Multiline = true;
            this.groupListBox.Name = "groupListBox";
            this.groupListBox.ReadOnly = true;
            this.groupListBox.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.groupListBox.Size = new System.Drawing.Size(452, 475);
            this.groupListBox.TabIndex = 3;
            // 
            // addGroupButton
            // 
            this.addGroupButton.Location = new System.Drawing.Point(491, 30);
            this.addGroupButton.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.addGroupButton.Name = "addGroupButton";
            this.addGroupButton.Size = new System.Drawing.Size(161, 28);
            this.addGroupButton.TabIndex = 4;
            this.addGroupButton.Text = "Добавить группу";
            this.addGroupButton.UseVisualStyleBackColor = true;
            this.addGroupButton.Click += new System.EventHandler(this.addGroupButton_Click);
            // 
            // removeGroupButton
            // 
            this.removeGroupButton.Location = new System.Drawing.Point(491, 65);
            this.removeGroupButton.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.removeGroupButton.Name = "removeGroupButton";
            this.removeGroupButton.Size = new System.Drawing.Size(161, 28);
            this.removeGroupButton.TabIndex = 5;
            this.removeGroupButton.Text = "Удалить группу";
            this.removeGroupButton.UseVisualStyleBackColor = true;
            this.removeGroupButton.Click += new System.EventHandler(this.removeGroupButton_Click);
            // 
            // timeTableButton
            // 
            this.timeTableButton.Location = new System.Drawing.Point(951, 15);
            this.timeTableButton.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.timeTableButton.Name = "timeTableButton";
            this.timeTableButton.Size = new System.Drawing.Size(100, 28);
            this.timeTableButton.TabIndex = 6;
            this.timeTableButton.Text = "Расписание";
            this.timeTableButton.UseVisualStyleBackColor = true;
            this.timeTableButton.Click += new System.EventHandler(this.timeTableButton_Click);
            // 
            // timeTableTextBox
            // 
            this.timeTableTextBox.Location = new System.Drawing.Point(477, 202);
            this.timeTableTextBox.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.timeTableTextBox.Multiline = true;
            this.timeTableTextBox.Name = "timeTableTextBox";
            this.timeTableTextBox.ReadOnly = true;
            this.timeTableTextBox.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.timeTableTextBox.Size = new System.Drawing.Size(400, 336);
            this.timeTableTextBox.TabIndex = 7;
            // 
            // writeTimetableButton
            // 
            this.writeTimetableButton.Location = new System.Drawing.Point(887, 122);
            this.writeTimetableButton.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.writeTimetableButton.Name = "writeTimetableButton";
            this.writeTimetableButton.Size = new System.Drawing.Size(164, 28);
            this.writeTimetableButton.TabIndex = 8;
            this.writeTimetableButton.Text = "Узнать расписание";
            this.writeTimetableButton.UseVisualStyleBackColor = true;
            this.writeTimetableButton.Click += new System.EventHandler(this.writeTimetableButton_Click);
            // 
            // changeGroupList
            // 
            this.changeGroupList.Location = new System.Drawing.Point(491, 101);
            this.changeGroupList.Margin = new System.Windows.Forms.Padding(4);
            this.changeGroupList.Name = "changeGroupList";
            this.changeGroupList.Size = new System.Drawing.Size(161, 28);
            this.changeGroupList.TabIndex = 9;
            this.changeGroupList.Text = "Изменить список";
            this.changeGroupList.UseVisualStyleBackColor = true;
            this.changeGroupList.Click += new System.EventHandler(this.changeGroupList_Click);
            // 
            // mainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1067, 554);
            this.Controls.Add(this.changeGroupList);
            this.Controls.Add(this.writeTimetableButton);
            this.Controls.Add(this.timeTableTextBox);
            this.Controls.Add(this.timeTableButton);
            this.Controls.Add(this.removeGroupButton);
            this.Controls.Add(this.addGroupButton);
            this.Controls.Add(this.groupListBox);
            this.Controls.Add(this.createTable);
            this.Controls.Add(this.groupsListBox);
            this.Controls.Add(this.button1);
            this.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.Name = "mainForm";
            this.Text = "Расписание групп";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.ComboBox groupsListBox;
        private System.Windows.Forms.Button createTable;
        private System.Windows.Forms.TextBox groupListBox;
        private System.Windows.Forms.Button addGroupButton;
        private System.Windows.Forms.Button removeGroupButton;
        private System.Windows.Forms.Button timeTableButton;
        private System.Windows.Forms.TextBox timeTableTextBox;
        private System.Windows.Forms.Button writeTimetableButton;
        private System.Windows.Forms.Button changeGroupList;
    }
}

