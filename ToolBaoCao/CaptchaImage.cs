﻿using System;
using System.Collections.Generic;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.Drawing;
using System.Linq;
using System.Web;

namespace ToolBaoCao
{
    namespace CaptchaImage
    {
        public class RandomImage
        {
            //Default Constructor
            public RandomImage() { }

            //property
            public string Text { get { return this.text; } }

            public Bitmap Image { get { return this.image; } }

            public int Width { get { return this.width; } }

            public int Height { get { return this.height; } }

            //Private variable
            private string text;

            private int width;
            private int height;
            private Bitmap image;
            private Random random = new Random();

            //Methods declaration
            public RandomImage(string s, int width, int height)
            {
                this.text = s;
                this.SetDimensions(width, height);
                this.GenerateImage();
            }

            public void Dispose()
            {
                GC.SuppressFinalize(this);
                this.Dispose(true);
            }

            protected virtual void Dispose(bool disposing)
            {
                if (disposing) this.image.Dispose();
            }

            private void SetDimensions(int width, int height)
            {
                if (width <= 0)
                    throw new ArgumentOutOfRangeException("width", width,
                        "Argument out of range, must be greater than zero.");
                if (height <= 0)
                    throw new ArgumentOutOfRangeException("height", height,
                        "Argument out of range, must be greater than zero.");
                this.width = width;
                this.height = height;
            }

            private void GenerateImage()
            {
                Bitmap bitmap = new Bitmap
                  (this.width, this.height, PixelFormat.Format32bppArgb);
                Graphics g = Graphics.FromImage(bitmap);
                g.SmoothingMode = SmoothingMode.AntiAlias;
                Rectangle rect = new Rectangle(0, 0, this.width, this.height);
                HatchBrush hatchBrush = new HatchBrush(HatchStyle.SmallConfetti,
                    Color.LightGray, Color.White);
                g.FillRectangle(hatchBrush, rect);
                SizeF size;
                float fontSize = rect.Height + 1;
                Font font;

                do
                {
                    fontSize--;
                    font = new Font(FontFamily.GenericSansSerif, fontSize, FontStyle.Bold);
                    size = g.MeasureString(this.text, font);
                } while (size.Width > rect.Width);
                StringFormat format = new StringFormat();
                format.Alignment = StringAlignment.Center;
                format.LineAlignment = StringAlignment.Center;
                GraphicsPath path = new GraphicsPath();
                //path.AddString(this.text, font.FontFamily, (int) font.Style,
                //    font.Size, rect, format);
                path.AddString(this.text, font.FontFamily, (int)font.Style, 75, rect, format);
                float v = 4F;
                PointF[] points =
                  {
                new PointF(this.random.Next(rect.Width) / v, this.random.Next( rect.Height) / v),
                new PointF(rect.Width - this.random.Next(rect.Width) / v, this.random.Next(rect.Height) / v),
                new PointF(this.random.Next(rect.Width) / v, rect.Height - this.random.Next(rect.Height) / v),
                new PointF(rect.Width - this.random.Next(rect.Width) / v, rect.Height - this.random.Next(rect.Height) / v)
          };
                Matrix matrix = new Matrix();
                matrix.Translate(0F, 0F);
                path.Warp(points, rect, matrix, WarpMode.Perspective, 0F);
                hatchBrush = new HatchBrush(HatchStyle.Percent10, Color.Black, Color.SkyBlue);
                g.FillPath(hatchBrush, path);
                int m = Math.Max(rect.Width, rect.Height);
                for (int i = 0; i < (int)(rect.Width * rect.Height / 30F); i++)
                {
                    int x = this.random.Next(rect.Width);
                    int y = this.random.Next(rect.Height);
                    int w = this.random.Next(m / 50);
                    int h = this.random.Next(m / 50);
                    g.FillEllipse(hatchBrush, x, y, w, h);
                }
                font.Dispose();
                hatchBrush.Dispose();
                g.Dispose();
                this.image = bitmap;
            }
        }

        public class Generate
        {
            public Generate()
            {
            }

            public static string RandomString(int stringlen)
            {
                Random r = new Random();
                if (stringlen <= 0) stringlen = 1;
                string s = "";

                for (int j = 0; j < stringlen; j++)
                {
                    int i = r.Next(3);
                    int ch;
                    switch (i)
                    {
                        case 1:
                            ch = r.Next(0, 9);
                            s += ch.ToString();
                            break;

                        case 2:
                            ch = r.Next(65, 90);
                            s += Convert.ToChar(ch).ToString();
                            break;

                        case 3:
                            ch = r.Next(97, 122);
                            s += Convert.ToChar(ch).ToString();
                            break;

                        default:
                            ch = r.Next(97, 122);
                            s += Convert.ToChar(ch).ToString();
                            break;
                    }
                    r.NextDouble();
                    r.Next(100, 1999);
                }
                return s;
            }

            public static string RandomString(int MinLen, int MaxLen)
            {
                Random r = new Random();
                if (MinLen <= 0) MinLen = 1;
                if (MaxLen < MinLen) MaxLen = MinLen;
                string s = "";

                for (int j = 0; j < r.Next(MinLen, MaxLen); j++)
                {
                    int i = r.Next(3);
                    int ch;
                    switch (i)
                    {
                        case 1:
                            ch = r.Next(0, 9);
                            s = s + ch.ToString();
                            break;

                        case 2:
                            ch = r.Next(65, 90);
                            s = s + Convert.ToChar(ch).ToString();
                            break;

                        case 3:
                            ch = r.Next(97, 122);
                            s = s + Convert.ToChar(ch).ToString();
                            break;

                        default:
                            ch = r.Next(97, 122);
                            s = s + Convert.ToChar(ch).ToString();
                            break;
                    }
                    r.NextDouble();
                    r.Next(100, 1999);
                }
                return s;
            }
        }
    }
}