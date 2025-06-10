using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.Windows.Forms;
using SkiaSharp;

namespace MissionPlanner.Controls
{
    public partial class VideoForm : Form
    {
        public VideoForm()
        {
            InitializeComponent();
        }

        public void UpdateFrame(Bitmap frame)
        {
            if (IsDisposed)
                return;

            if (InvokeRequired)
            {
                BeginInvoke(new Action<Bitmap>(UpdateFrame), frame);
                return;
            }

            var old = pictureBox1.Image;
            if (frame == null)
            {
                pictureBox1.Image = null;
            }
            else
            {
                pictureBox1.Image = new Bitmap(frame.Width, frame.Height, 4 * frame.Width,
                    System.Drawing.Imaging.PixelFormat.Format32bppPArgb,
                    frame.LockBits(Rectangle.Empty, null, SKColorType.Bgra8888).Scan0);
            }
            old?.Dispose();
        }

        private void VideoForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            e.Cancel = true;
            Hide();
        }
    }
}
