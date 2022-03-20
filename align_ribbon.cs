using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace AlignTool
{
    public partial class align_ribbon
    {
        bool is_to_shape = true;

        private void align_ribbon_Load(object sender, RibbonUIEventArgs e)
        {
            update_enable();
        }

        private void update_enable()
        {
            to_shape.Enabled = !is_to_shape;
            to_slide.Enabled = is_to_shape;
        }

        private void left_Click(object sender, RibbonControlEventArgs e)
        {
            align_shapes("left");
        }

        private void horizontal_Click(object sender, RibbonControlEventArgs e)
        {
            align_shapes("horizontal");

        }

        private void right_Click(object sender, RibbonControlEventArgs e)
        {
            align_shapes("right");

        }

        private void top_Click(object sender, RibbonControlEventArgs e)
        {
            align_shapes("top");
        }

        private void vertical_Click(object sender, RibbonControlEventArgs e)
        {
            align_shapes("vertical");
        }

        private void bottom_Click(object sender, RibbonControlEventArgs e)
        {
            align_shapes("bottom");

        }

        private void hori_dist_Click(object sender, RibbonControlEventArgs e)
        {
            align_shapes("hori_dist");

        }

        private void vert_dist_Click(object sender, RibbonControlEventArgs e)
        {
            align_shapes("vert_dist");

        }

        private void hori_group_Click(object sender, RibbonControlEventArgs e)
        {
            align_shapes("hori_group");

        }
        private void vert_group_Click(object sender, RibbonControlEventArgs e)
        {
            align_shapes("vert_group");
        }

        private void to_slide_Click(object sender, RibbonControlEventArgs e)
        {
            is_to_shape = false;
            update_enable();
        }

        private void to_shape_Click(object sender, RibbonControlEventArgs e)
        {
            is_to_shape = true;
            update_enable();
        }

        private void copy_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (Globals.ThisAddIn.Application.ActiveWindow.Selection.ShapeRange.Count > 0)
                {
                    var shape = Globals.ThisAddIn.Application.ActiveWindow.Selection.ShapeRange;

                    int count = shape.Count;


                    var shapeRange = Globals.ThisAddIn.Application.ActiveWindow.Selection.ShapeRange;

                    foreach (PowerPoint.Shape shp in shapeRange)
                    {
                        shp.Copy();
                        var new_shp = Globals.ThisAddIn.Application.ActiveWindow.View.Slide.Shapes.Paste();

                        new_shp.Top = shp.Top;
                        new_shp.Left = shp.Left;
                    }
                }
            }
            catch (Exception)
            {

            }
        }


        private void align_shapes(string align_type)
        {
            try
            {
                if (Globals.ThisAddIn.Application.ActiveWindow.Selection.ShapeRange.Count > 0)
                {
                    var shape = Globals.ThisAddIn.Application.ActiveWindow.Selection.ShapeRange;
                    
                    int count = shape.Count;

                    bool need_slide = false;

                    if (count == 1)
                    {
                        if (align_type == "hori_dist" || align_type == "vert_dist") { return; }
                    }

                    if (!this.is_to_shape || count == 1) { need_slide = true; }

                    if (need_slide) { count += 1; }

                    int[] index = new int[count];
                    for (int j = 0; j < count; j++)
                    {
                        index[j] = j;
                    }

                    float[] centers_x = new float[count];
                    float[] centers_y = new float[count];
                    float[,] sizes = new float[count, 2];
                    float[,] corners = new float[count, 2];

                    float left_top_x = 99999;
                    float left_top_y = 99999;
                    float right_bottom_x = 0;
                    float right_bottom_y = 0;

                    var shapeRange = Globals.ThisAddIn.Application.ActiveWindow.Selection.ShapeRange;
                    int i = 0;
                    foreach (PowerPoint.Shape shp in shapeRange)
                    {

                        float w = shp.Width;
                        float h = shp.Height;
                        float t = shp.Top;
                        float l = shp.Left;

                        centers_x[i] = l + w / 2;
                        centers_y[i] = t + h / 2;

                        sizes[i, 0] = w;
                        sizes[i, 1] = h;

                        corners[i, 0] = l;
                        corners[i, 1] = t;

                        float r = l + w;
                        float b = t + h;

                        if (left_top_x > l) left_top_x = l;
                        if (left_top_y > t) left_top_y = t;
                        if (right_bottom_x < r) right_bottom_x = r;
                        if (right_bottom_y < b) right_bottom_y = b;

                        i++;
                    }

                    i = count - 1;
                    if (need_slide)
                    {
                        float w = Globals.ThisAddIn.Application.ActivePresentation.PageSetup.SlideWidth;
                        float h = Globals.ThisAddIn.Application.ActivePresentation.PageSetup.SlideHeight;
                        float l = 0;
                        float t = 0;

                        centers_x[i] = l + w / 2;
                        centers_y[i] = t + h / 2;

                        sizes[i, 0] = w;
                        sizes[i, 1] = h;

                        corners[i, 0] = l;
                        corners[i, 1] = t;
                    }

                    // 开始对齐

                    if (align_type == "hori_dist" )
                    {
                        Array.Sort(centers_x, index);

                        int left_id = index[0];
                        int right_id = index[count-1];

                        float left_w = sizes[left_id, 0];
                        float right_w = sizes[right_id, 0];

                        float between_gap = (centers_x[count-1] - right_w / 2) - (centers_x[0] + left_w / 2) ;
                        float total_len = 0;
                        foreach (var id in index)
                        {
                            if (id == left_id || id == right_id) continue;
                            total_len += sizes[id, 0];
                        }
                        between_gap = between_gap - total_len;

                        between_gap /= count - 1;

                        float pre_right = centers_x[0] + left_w / 2;

                        foreach (var id in index)
                        {
                            if (id == left_id || id == right_id) continue;

                            int k = 0;
                            foreach (PowerPoint.Shape shp in shapeRange)
                            {

                                if (id == k) { 
                                    float shp_w = shp.Width;

                                    pre_right += between_gap;
                                    shp.Left = pre_right;

                                    pre_right += shp_w;
                                }
                                k++;
                            }
                        }

                    }
                    else if (align_type == "vert_dist")
                    {
                        Array.Sort(centers_y, index);

                        int top_id = index[0];
                        int bottom_id = index[count - 1];

                        float top_h = sizes[top_id, 1];
                        float bottom_h = sizes[bottom_id, 1];

                        float between_gap = (centers_y[count - 1] - bottom_h / 2) - (centers_y[0] + top_h / 2);
                        float total_len = 0;
                        foreach (var id in index)
                        {
                            if (id == top_id || id == bottom_id) continue;
                            total_len += sizes[id, 1];
                        }
                        between_gap = between_gap - total_len;

                        between_gap /= count - 1;

                        float pre_botttom = centers_y[0] + top_h / 2;

                        foreach (var id in index)
                        {
                            if (id == top_id || id == bottom_id) continue;

                            int k = 0;
                            foreach (PowerPoint.Shape shp in shapeRange)
                            {

                                if (id == k)
                                {
                                    float shp_h = shp.Height;

                                    pre_botttom += between_gap;
                                    shp.Top = pre_botttom;

                                    pre_botttom += shp_h;
                                }
                                k++;
                            }
                        }
                    }
                    else if (align_type == "hori_group")
                    {
                        float group_center_y = (left_top_y + right_bottom_y) / 2;

                        float slide_h = Globals.ThisAddIn.Application.ActivePresentation.PageSetup.SlideHeight;

                        float slide_center_y = slide_h / 2;

                        float slide_y_offset = slide_center_y - group_center_y;

                        foreach (PowerPoint.Shape shp in shapeRange)
                        {
                            shp.Top += slide_y_offset;
                        }
                    }
                    else if (align_type == "vert_group")
                    {
                        float group_center_x = (left_top_x + right_bottom_x) / 2;
                        
                        float slide_w = Globals.ThisAddIn.Application.ActivePresentation.PageSetup.SlideWidth;
                        
                        float slide_center_x = slide_w / 2;

                        float slide_x_offset = slide_center_x - group_center_x;

                        foreach (PowerPoint.Shape shp in shapeRange)
                        {
                            shp.Left += slide_x_offset;
                        }
                    }
                    else
                    {
                        i = 0;
                        foreach (PowerPoint.Shape shp in shapeRange)
                        {
                            switch (align_type)
                            {
                                case "horizontal":
                                    shp.Top = centers_y[count - 1] - sizes[i, 1] / 2;
                                    break;
                                case "vertical":
                                    shp.Left = centers_x[count - 1] - sizes[i, 0] / 2;
                                    break;

                                case "left":
                                    shp.Left = corners[count - 1, 0];
                                    break;
                                case "right":
                                    shp.Left = (corners[count - 1, 0] + sizes[count - 1, 0] - sizes[i, 0]);
                                    break;

                                case "top":
                                    shp.Top = corners[count - 1, 1];
                                    break;
                                case "bottom":
                                    shp.Top = (corners[count - 1, 1] + sizes[count-1, 1] - sizes[i, 1] );
                                    break;

                                default:
                                    break;
                            }
                            i++;
                        }
                    }
                }

            }
            catch (Exception)
            {

            }

        }
    }
}
