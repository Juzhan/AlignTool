
namespace AlignTool
{
    partial class align_ribbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public align_ribbon()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region 组件设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            this.tab1 = this.Factory.CreateRibbonTab();
            this.tool_group = this.Factory.CreateRibbonGroup();
            this.align_menu = this.Factory.CreateRibbonMenu();
            this.left = this.Factory.CreateRibbonButton();
            this.horizontal = this.Factory.CreateRibbonButton();
            this.right = this.Factory.CreateRibbonButton();
            this.separator1 = this.Factory.CreateRibbonSeparator();
            this.top = this.Factory.CreateRibbonButton();
            this.vertical = this.Factory.CreateRibbonButton();
            this.bottom = this.Factory.CreateRibbonButton();
            this.separator2 = this.Factory.CreateRibbonSeparator();
            this.hori_dist = this.Factory.CreateRibbonButton();
            this.vert_dist = this.Factory.CreateRibbonButton();
            this.separator3 = this.Factory.CreateRibbonSeparator();
            this.hori_group = this.Factory.CreateRibbonButton();
            this.vert_group = this.Factory.CreateRibbonButton();
            this.separator4 = this.Factory.CreateRibbonSeparator();
            this.to_slide = this.Factory.CreateRibbonButton();
            this.to_shape = this.Factory.CreateRibbonButton();
            this.copy = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.tool_group.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.ControlId.OfficeId = "TabHome";
            this.tab1.Groups.Add(this.tool_group);
            this.tab1.Label = "TabHome";
            this.tab1.Name = "tab1";
            // 
            // tool_group
            // 
            this.tool_group.Items.Add(this.align_menu);
            this.tool_group.Items.Add(this.copy);
            this.tool_group.Label = "Tools";
            this.tool_group.Name = "tool_group";
            // 
            // align_menu
            // 
            this.align_menu.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.align_menu.Description = "在选择的形状集合中，以最后一个形状为基准对齐";
            this.align_menu.Image = global::AlignTool.Properties.Resources.align;
            this.align_menu.Items.Add(this.left);
            this.align_menu.Items.Add(this.horizontal);
            this.align_menu.Items.Add(this.right);
            this.align_menu.Items.Add(this.separator1);
            this.align_menu.Items.Add(this.top);
            this.align_menu.Items.Add(this.vertical);
            this.align_menu.Items.Add(this.bottom);
            this.align_menu.Items.Add(this.separator2);
            this.align_menu.Items.Add(this.hori_dist);
            this.align_menu.Items.Add(this.vert_dist);
            this.align_menu.Items.Add(this.separator3);
            this.align_menu.Items.Add(this.hori_group);
            this.align_menu.Items.Add(this.vert_group);
            this.align_menu.Items.Add(this.separator4);
            this.align_menu.Items.Add(this.to_slide);
            this.align_menu.Items.Add(this.to_shape);
            this.align_menu.Label = "对齐";
            this.align_menu.Name = "align_menu";
            this.align_menu.ShowImage = true;
            this.align_menu.SuperTip = "在选择的形状集合中，以最后一个形状为基准对齐";
            // 
            // left
            // 
            this.left.Image = global::AlignTool.Properties.Resources.left;
            this.left.Label = "左对齐";
            this.left.Name = "left";
            this.left.ShowImage = true;
            this.left.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.left_Click);
            // 
            // horizontal
            // 
            this.horizontal.Image = global::AlignTool.Properties.Resources.hori;
            this.horizontal.Label = "横向居中";
            this.horizontal.Name = "horizontal";
            this.horizontal.ShowImage = true;
            this.horizontal.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.horizontal_Click);
            // 
            // right
            // 
            this.right.Image = global::AlignTool.Properties.Resources.right;
            this.right.Label = "右对齐";
            this.right.Name = "right";
            this.right.ShowImage = true;
            this.right.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.right_Click);
            // 
            // separator1
            // 
            this.separator1.Name = "separator1";
            // 
            // top
            // 
            this.top.Image = global::AlignTool.Properties.Resources.top;
            this.top.Label = "顶对齐";
            this.top.Name = "top";
            this.top.ShowImage = true;
            this.top.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.top_Click);
            // 
            // vertical
            // 
            this.vertical.Image = global::AlignTool.Properties.Resources.vert;
            this.vertical.Label = "纵向对齐";
            this.vertical.Name = "vertical";
            this.vertical.ShowImage = true;
            this.vertical.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.vertical_Click);
            // 
            // bottom
            // 
            this.bottom.Image = global::AlignTool.Properties.Resources.bottom;
            this.bottom.Label = "底对齐";
            this.bottom.Name = "bottom";
            this.bottom.ShowImage = true;
            this.bottom.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.bottom_Click);
            // 
            // separator2
            // 
            this.separator2.Name = "separator2";
            // 
            // hori_dist
            // 
            this.hori_dist.Image = global::AlignTool.Properties.Resources.hori_dist;
            this.hori_dist.Label = "横向等距";
            this.hori_dist.Name = "hori_dist";
            this.hori_dist.ShowImage = true;
            this.hori_dist.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.hori_dist_Click);
            // 
            // vert_dist
            // 
            this.vert_dist.Image = global::AlignTool.Properties.Resources.vert_dist;
            this.vert_dist.Label = "纵向等距";
            this.vert_dist.Name = "vert_dist";
            this.vert_dist.ShowImage = true;
            this.vert_dist.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.vert_dist_Click);
            // 
            // separator3
            // 
            this.separator3.Name = "separator3";
            // 
            // hori_group
            // 
            this.hori_group.Image = global::AlignTool.Properties.Resources.hori_group;
            this.hori_group.Label = "横向居中（集合）";
            this.hori_group.Name = "hori_group";
            this.hori_group.ShowImage = true;
            this.hori_group.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.hori_group_Click);
            // 
            // vert_group
            // 
            this.vert_group.Image = global::AlignTool.Properties.Resources.vert_group;
            this.vert_group.Label = "纵向居中（集合）";
            this.vert_group.Name = "vert_group";
            this.vert_group.ShowImage = true;
            this.vert_group.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.vert_group_Click);
            // 
            // separator4
            // 
            this.separator4.Name = "separator4";
            // 
            // to_slide
            // 
            this.to_slide.Label = "与幻灯片对齐";
            this.to_slide.Name = "to_slide";
            this.to_slide.ShowImage = true;
            this.to_slide.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.to_slide_Click);
            // 
            // to_shape
            // 
            this.to_shape.Label = "与形状对齐";
            this.to_shape.Name = "to_shape";
            this.to_shape.ShowImage = true;
            this.to_shape.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.to_shape_Click);
            // 
            // copy
            // 
            this.copy.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.copy.Description = "复制形状，并且和被复制形状位置一样";
            this.copy.Image = global::AlignTool.Properties.Resources.copy;
            this.copy.Label = "原位复制";
            this.copy.Name = "copy";
            this.copy.ShowImage = true;
            this.copy.SuperTip = "复制新的形状，并且和被复制形状位置一样";
            this.copy.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.copy_Click);
            // 
            // align_ribbon
            // 
            this.Name = "align_ribbon";
            this.RibbonType = "Microsoft.PowerPoint.Presentation";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.align_ribbon_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.tool_group.ResumeLayout(false);
            this.tool_group.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup tool_group;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu align_menu;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton left;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton horizontal;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton right;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton top;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton vertical;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton bottom;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton hori_dist;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton vert_dist;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator3;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton to_shape;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton to_slide;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton copy;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton hori_group;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator4;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton vert_group;
    }

    partial class ThisRibbonCollection
    {
        internal align_ribbon align_ribbon
        {
            get { return this.GetRibbon<align_ribbon>(); }
        }
    }
}
