﻿using System.Windows.Forms;
using System.Threading;
using System.Diagnostics;

namespace WaitDialogForm
{
    /// <summary>
    /// WaitDialog の概要の説明です。
    /// </summary>
    public class WaitDialog : System.Windows.Forms.Form
    {
        private bool bAborting = false;     // 中止フラグ
        private bool bShowing = false;      // ダイアログ表示中フラグ

        public System.Windows.Forms.Label labelMsg;
        public System.Windows.Forms.ProgressBar progBarMeter;
        private System.Windows.Forms.Button btnCancel;

        /// <summary>
        /// 必要なデザイナ変数です。
        /// </summary>
        private System.ComponentModel.Container components = null;

        public WaitDialog()
        {
            //
            // Windows フォーム デザイナ サポートに必要です。
            //
            InitializeComponent();

            //
            // TODO: InitializeComponent 呼び出しの後に、コンストラクタ コードを追加してください。
            //
            this.Title = "画像挿入中 ...";
            this.Msg = "処理準備中、しばらくお待ちください";
            this.ProgressMin   = 0;  // 処理件数の最小値（0件から開始）
            this.ProgressStep  = 1;  // 何件ごとにメーターを進めるか
            this.ProgressValue = 0;  // 最初の件数
        }

        /// <summary>
        /// 使用されているリソースに後処理を実行します。
        /// </summary>
        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                if (components != null)
                {
                    components.Dispose();
                }
            }
            base.Dispose(disposing);
        }

        #region Windows フォーム デザイナで生成されたコード 
        /// <summary>
        /// デザイナ サポートに必要なメソッドです。このメソッドの内容を
        /// コード エディタで変更しないでください。
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(WaitDialog));
            this.labelMsg = new System.Windows.Forms.Label();
            this.progBarMeter = new System.Windows.Forms.ProgressBar();
            this.btnCancel = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // labelMsg
            // 
            this.labelMsg.Location = new System.Drawing.Point(18, 9);
            this.labelMsg.Name = "labelMsg";
            this.labelMsg.Size = new System.Drawing.Size(408, 16);
            this.labelMsg.TabIndex = 0;
            this.labelMsg.Text = "処理準備中、しばらくお待ちください";
            // 
            // progBarMeter
            // 
            this.progBarMeter.Location = new System.Drawing.Point(18, 35);
            this.progBarMeter.Name = "progBarMeter";
            this.progBarMeter.Size = new System.Drawing.Size(408, 23);
            this.progBarMeter.TabIndex = 3;
            // 
            // btnCancel
            // 
            this.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnCancel.Location = new System.Drawing.Point(179, 64);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(88, 23);
            this.btnCancel.TabIndex = 4;
            this.btnCancel.Text = "キャンセル";
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // WaitDialog
            // 
            this.AutoScaleBaseSize = new System.Drawing.Size(7, 18);
            this.CancelButton = this.btnCancel;
            this.ClientSize = new System.Drawing.Size(440, 95);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.progBarMeter);
            this.Controls.Add(this.labelMsg);
            this.Font = new System.Drawing.Font("メイリオ", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "WaitDialog";
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "画像挿入中 ...";
            this.Closing += new System.ComponentModel.CancelEventHandler(this.WaitDialog_Closing);
            this.ResumeLayout(false);

        }
        #endregion

        /// <summary>
        /// ShowDialogメソッドのシャドウ（WaitDialogクラスでは、ShowDialogメソッドは使用不可）
        /// </summary>
        public new DialogResult ShowDialog()
        {
            Debug.Assert(false,
              "WaitDialogクラスはShowDialogメソッドを利用できません。\n" +
              "Showメソッドを使ってモードレス・ダイアログを構築してください。");
            return DialogResult.Abort;
        }

        /// <summary>
        /// Showメソッドのシャドウ（シャドウ＝new修飾子）
        /// </summary>
        public new void Show()
        {
            // 変数の初期化
            this.DialogResult = DialogResult.OK;
            this.bAborting = false;

            base.Show();
            this.bShowing = true;
        }

        /// <summary>
        /// Closeメソッドのシャドウ
        /// </summary>
        public new void Close()
        {
            this.bShowing = false;
            base.Close();
        }

        /// <summary>
        /// キャンセル・ボタンが押されたときの処理
        /// </summary>
        /// <remarks>処理を途中でキャンセル（中断）する。</remarks>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnCancel_Click(object sender, System.EventArgs e)
        {
            Abort();    // 中止処理
        }

        /// <summary>
        /// 中止（キャンセル）処理
        /// </summary>
        private void Abort()
        {
            this.bAborting = true;
            this.DialogResult = DialogResult.Abort;
        }

        /// <summary>
        /// ダイアログが閉じられるときの処理
        /// </summary>
        /// <remarks>
        /// 右上の［閉じる］ボタンが押された場合には、
        /// ［キャンセル］ボタンと同じように、処理を途中でキャンセル（中断）する。
        /// </remarks>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void WaitDialog_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            if (bShowing == true)
            {
                // ダイアログ表示中なので中止（キャンセル）処理を実行
                Abort();
                // まだダイアログは閉じない
                e.Cancel = true;
            }
            else
            {
                // フォームは閉じられるところので素直に閉じる
                e.Cancel = false;
            }
        }

        /// **************************************************************

        /// <summary>
        /// 処理がキャンセル（中止）されているかどうかを示す値を取得する。
        /// キャンセルされた場合はtrue。それ以外はfalse。
        /// </summary>
        public bool IsAborting
        {
            get
            {
                return this.bAborting;
            }
        }

        /// <summary>
        /// ダイアログのテキストを設定する。
        /// </summary>
        /// <remarks>
        /// 処理の概要を表示する。
        /// 例えば、画像表示機能であれば、「画像表示」のように表示する。
        /// </remarks>
        public string Title
        {
            set
            {
                this.Text = value;
            }
        }

        /// <summary>
        /// メイン・メッセージのテキストを設定する。
        /// </summary>
        /// <remarks>
        /// 処理の概要を表示する。
        /// 例えば、ファイルの転送を行っているなら、「ファイルを転送しています……」のように表示する。
        /// </remarks>
        public string Msg
        {
            set
            {
                this.labelMsg.Text = value;
            }
        }

        /// <summary>
        /// 進行状況メーターの現在位置を設定する。
        /// </summary>
        /// <remarks>
        /// 例えば、処理に200の工数があった場合、現在その200の工数の中のどの位置にいるかを示す値。
        /// 既定値は「0」。
        /// </remarks>
        public int ProgressValue
        {
            set
            {
                this.progBarMeter.Value = value;
            }
        }

        /// <summary>
        /// 進行状況メーターの範囲の最大値を設定する。
        /// </summary>
        /// <remarks>
        /// 処理に200の工数があるなら「200」になる。
        /// 既定値は「100」。
        /// </remarks>
        public int ProgressMax
        {
            set
            {
                this.progBarMeter.Maximum = value;
            }
        }

        /// <summary>
        /// 進行状況メーターの範囲の最小値を設定する。
        /// </summary>
        /// <remarks>
        /// 既定値は「0」。
        /// </remarks>
        public int ProgressMin
        {
            set
            {
                this.progBarMeter.Minimum = value;
            }
        }

        /// <summary>
        /// PerformStepメソッドを呼び出したときに、進行状況メーターの現在位置を進める量（ProgressValueの増分値）を設定する。
        /// </summary>
        /// <remarks>
        /// 処理工数が200で、5つの工数が終わった段階で進行状況メーターを更新したい場合は「5」にする。
        /// 既定値は「10」。
        /// </remarks>
        public int ProgressStep
        {
            set
            {
                this.progBarMeter.Step = value;
            }
        }

        /// <summary>
        /// 進行状況メーターの現在位置（ProgressValue）をProgressStepプロパティの量だけ進める。
        /// </summary>
        public void PerformStep()
        {
            this.progBarMeter.PerformStep();
        }
    }
}