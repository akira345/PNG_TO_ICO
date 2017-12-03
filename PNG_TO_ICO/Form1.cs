using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace PNG_TO_ICO
{
    public partial class Form1 : Form
    {
        private string save_base_path;

        //DllImport属性を用いて定義したメソッドは、NativeMethodsというクラスに含めるという規則になっているらしい
        internal static class NativeMethods
        {
            [DllImport("user32.dll", CharSet = CharSet.Auto)]
            public static extern bool DestroyIcon(IntPtr handle);
        }

        public Form1()
        {
            InitializeComponent();
            AllowDrop = true; //D&D許可
            DragDrop += new DragEventHandler(Form1_DragDrop);
            DragEnter += new DragEventHandler(Form1_DragEnter);
            //大きさ固定
            this.FormBorderStyle = FormBorderStyle.FixedSingle;
            //フォームが最大化されないようにする
            this.MaximizeBox = false;
            //フォームが最小化されないようにする
            this.MinimizeBox = false;
            //カレントディレクトリを初期値にする。
            save_base_path = Environment.CurrentDirectory;

        }
        private void Form1_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
                e.Effect = DragDropEffects.Copy;
            else
                e.Effect = DragDropEffects.None;
        }
        private void Form1_DragDrop(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
                try
                {
                    //https://dobon.net/vb/dotnet/graphics/imagefromfile.html より
                    //画像ファイルを読み込んで、Imageオブジェクトを作成する
                    Image img = Image.FromFile(@files[0]);


                    Bitmap bmp = new Bitmap(img);

                    //縮小
                    int resizeWidth = 128;
                    int resizeHeight = (int)(bmp.Height * ((double)resizeWidth / (double)bmp.Width));
                    Bitmap resizeBmp = new Bitmap(resizeWidth, resizeHeight);
                    Graphics g = Graphics.FromImage(resizeBmp);
                    g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                    //g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.NearestNeighbor;
                    g.DrawImage(bmp, 0, 0, resizeWidth, resizeHeight);
                    g.Dispose();

                    //Debug
                    //ピクチャボックスとフォームを画像めいっぱいのサイズに広げる
                    this.Size = new System.Drawing.Size(256, 256);
                    pictureBox1.Size = new System.Drawing.Size(resizeBmp.Width, resizeBmp.Height);
                    


                    pictureBox1.Image = resizeBmp;

                    //BitmapからIconを作成
                    Icon ico = Icon.FromHandle(resizeBmp.GetHicon());

                    

                    //ダイアログ表示
                    //SaveFileDialogクラスのインスタンスを作成
                    SaveFileDialog sfd = new SaveFileDialog();

                    //はじめのファイル名を指定する
                    //はじめに「ファイル名」で表示される文字列を指定する
                    sfd.FileName = files[0] + ".ico";
                    //はじめに表示されるフォルダを指定する
                    sfd.InitialDirectory = save_base_path;
                    //[ファイルの種類]に表示される選択肢を指定する
                    //指定しない（空の文字列）の時は、現在のディレクトリが表示される
                    sfd.Filter = "icoファイル(*.ico)|*.ico";
                    //[ファイルの種類]ではじめに選択されるものを指定する
                    //タイトルを設定する
                    sfd.Title = "保存先のファイルを選択してください";
                    //ダイアログボックスを閉じる前に現在のディレクトリを復元するようにする
                    sfd.RestoreDirectory = true;
                    //既に存在するファイル名を指定したとき警告する
                    //デフォルトでTrueなので指定する必要はない
                    sfd.OverwritePrompt = true;
                    //存在しないパスが指定されたとき警告を表示する
                    //デフォルトでTrueなので指定する必要はない
                    sfd.CheckPathExists = true;

                    //ダイアログを表示する
                    if (sfd.ShowDialog() == DialogResult.OK)
                    {
                        //OKボタンがクリックされたとき、選択されたファイル名を表示する
                        //ファイル削除し保存する。
                        File.Delete(sfd.FileName);

                        //書き込む
                        //https://dobon.net/vb/dotnet/graphics/saveimage.html 参照
                        FileStream fs = new FileStream(
                            sfd.FileName,
                            FileMode.Create,
                            FileAccess.Write);
                        ico.Save(fs);
                        fs.Close();
                        //後片付け
                        ico.Dispose();
                        bmp.Dispose();
                        //Icon.FromHandleでIconを作成したときはDestroyIconで破棄する
                        NativeMethods.DestroyIcon(ico.Handle);
                    }
                }
                catch (Exception)
                {
                    MessageBox.Show("変換できません！！");
                    return;
                }

            }
        }

    }
}
