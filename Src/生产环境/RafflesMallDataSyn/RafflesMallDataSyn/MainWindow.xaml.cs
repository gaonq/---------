using RafflesMallDataSyn.WebReference;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Windows.Threading;
using Dapper;
using DapperExtensions;
using System.Threading;
using System.Runtime.InteropServices;
using System.IO;
using System.Xml.Serialization;
namespace RafflesMallDataSyn
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {
        DispatcherTimer timer = new DispatcherTimer();
        string ip;
        string storeCode;
        string mallid;
        string username;
        string password;
        string licensekey;
        string connString = ConfigurationSettings.AppSettings["ConnectionString"];
        List<Order> orderList;
        List<OrderItem> orderItemList;
        List<Order> orderRejectList;
        List<OrderItem> orderRejectItemList;
        string itemcode;
        public MainWindow()
        {
            InitializeComponent();
        }
        void MainWindow_Loaded(object sender, RoutedEventArgs e)
        {

            init();

            if (GetMacAddressBySendARP() != "00-E0-B4-17-C1-FD")
            {
                MessageBox.Show("此电脑未授权，你无权使用该系统");
                this.btnDataSync.Content = "此电脑未授权，你无权使用该系统";
                this.btnDataSync.IsEnabled = false;
            }
        }

        void init()
        {
            username = System.Configuration.ConfigurationSettings.AppSettings["username"].ToString(); ;
            storeCode = System.Configuration.ConfigurationSettings.AppSettings["storeCode"].ToString(); ;
            mallid = System.Configuration.ConfigurationSettings.AppSettings["mallid"].ToString(); ;
            password = System.Configuration.ConfigurationSettings.AppSettings["password"].ToString(); ;
            licensekey = System.Configuration.ConfigurationSettings.AppSettings["licensekey"].ToString(); ; ;
            itemcode = System.Configuration.ConfigurationSettings.AppSettings["itemcode"].ToString(); ; ;
            ip = string.Concat("      IP：  ", GetIP() + "    MAC： " + GetMacAddressBySendARP()); ;
            timer.Tick += new EventHandler(timer_Tick);
            timer.Interval = TimeSpan.FromSeconds(0.1);   //设置刷新的间隔时间
            timer.Start();
            this.StoreCodeTextBlock.Text = string.Concat(this.StoreCodeTextBlock.Text, storeCode);
            this.mallidTextBlock.Text = string.Concat(this.mallidTextBlock.Text, mallid);

        }
        /// <summary>
        /// 销售订单
        /// </summary>
        /// <returns></returns>
        public List<Order> GetOrderList()
        {
            List<Order> OrderList;
            try
            {
                using (var conn = new System.Data.SqlClient.SqlConnection(connString))
                {
                    conn.Open(); 
                    OrderList = conn.Query<Order>("select    * from CsmMaster as a  where   a.OptTime>='" + this.DateSelect.Text + "' and a.opttime<='" + this.DateSelect.Text + " 23:59:59' order by a.OptTime desc").ToList();
                    conn.Close();
                }

                return OrderList.ToList();
            }
            catch (Exception ex)
            {
                MessageBox.Show("获取数据失败:" + ex.Message);
                return null;
            }
        }


        public List<OrderItem> GetOrderItemList()
        {
            List<OrderItem> ItemList;
            try
            {
                using (var conn = new System.Data.SqlClient.SqlConnection(connString))
                {
                    conn.Open();
                    ItemList = conn.Query<OrderItem>("select   * from CsmDetail as a  where a.OptTime>='" + this.DateSelect.Text + "'  and a.opttime<='" + this.DateSelect.Text + " 23:59:59'").ToList();
                    conn.Close();
                }

                return ItemList.ToList();
            }
            catch (Exception ex)
            {
                MessageBox.Show("获取数据失败:" + ex.Message);
                return null;
            }
        }
        /// <summary>
        /// 退回订单
        /// </summary>
        /// <returns></returns>
        public List<Order> GetRejectOrderList()
        {
            List<Order> RejectOrderList;
            try
            {
                using (var conn = new System.Data.SqlClient.SqlConnection(connString))
                {
                    conn.Open();
                    RejectOrderList = conn.Query<Order>("select  * from RfdMaster as a  where a.OptTime>='" + this.DateSelect.Text + "' and a.opttime<='" + this.DateSelect.Text + " 23:59:59' order by a.OptTime desc").ToList();
                    conn.Close();
                }

                return RejectOrderList.ToList();
            }
            catch (Exception ex)
            {
                MessageBox.Show("获取数据失败:" + ex.Message);
                return null;
            }
        }

        public Order GetRejectOrder(string id)
        {
             Order order;
            try
            {
                using (var conn = new System.Data.SqlClient.SqlConnection(connString))
                {
                    conn.Open();
                    order = conn.Query<Order>("select   * from CsmMaster as a  where  a.exchno='" + id + "'").FirstOrDefault();
                    conn.Close();
                }

                return order;
            }
            catch (Exception ex)
            {
                MessageBox.Show("获取数据失败:" + ex.Message);
                return null;
            }
        }
        public List<OrderItem> GetRejectOrderItemList()
        {
            List<OrderItem> ItemList;
            try
            {
                using (var conn = new System.Data.SqlClient.SqlConnection(connString))
                {
                    conn.Open();
                    ItemList = conn.Query<OrderItem>("select   * from RfdDetail as a  where a.OptTime>='" + this.DateSelect.Text + "'  and a.opttime<='" + this.DateSelect.Text + " 23:59:59'").ToList();
                    conn.Close();
                }

                return ItemList.ToList();
            }
            catch (Exception ex)
            {
                MessageBox.Show("获取数据失败:" + ex.Message);
                return null;
            }
        }
        /// <summary>
        /// 提交销售的数据
        /// </summary>
        /// <param name="orderList"></param>
        /// <param name="itemList"></param>
        public void SendData(List<Order> orderList, List<OrderItem> itemList)
        {
            try
            {
            int k = 0;
            decimal sumPrice= 0;
            int sumCount = 0;
            foreach (Order order in orderList)
            {

                WebReference.sales sales = new WebReference.sales();

                WebReference.postsalescreaterequest request = new WebReference.postsalescreaterequest();

                #region header;//标头
                requestheader header = new requestheader();
                header.lang = ""; //语言
                header.pageno = 0;//页数 
                header.pagerecords = 0;//每页记录数
                header.updatecount = 0;//每次更新记录数 

                header.licensekey = licensekey;
                header.password = password;
                header.username = username;
                header.messageid = "332";//固定值：332
                header.messagetype = "SALESDATA";//固定值：SALESDATA
                header.version = "V332M";//固定值：
                request.header = header;
                #endregion

                sumPrice += order.TtlPrice1;

                #region 销售开单主表
                saleshdr salestotal = new saleshdr();
                // 销售开单主表 
                #region 可以为空
                //salestotal.buyerremark = "";//卖家备注 允许为空
                //salestotal.coupongroup = "";//优惠券组 允许为空
                //salestotal.couponnumber = "";//优惠券号码 允许为空
                //salestotal.couponqty = "";//优惠券数量 允许为空
                //salestotal.coupontype = "";//优惠券类型 允许为空
                //salestotal.demographiccode = "";//顾客统计代码 允许为空
                //salestotal.demographicdata = "";//顾客统计值 允许为空
                //salestotal.ecorderno = "";//网购订单号 允许为空
                //salestotal.extendparam = "";//扩展参数 允许为空
                //salestotal.invoicecontent = "";//发票内容 允许为空
                //salestotal.invoicetitle = "";//发票抬头 允许为空
                //salestotal.orderremark = "";//交易备注 允许为空
                //salestotal.orgstorecode = "";//原交易店铺号 允许为空 退货时，原交易店铺号 如果是按单退货，提供此店铺号
                //salestotal.orgtillid = "";// 原收银机号 允许为空 退货时，原收银机号  如果是按单退货，提供此收银机号
                //salestotal.shoptaxgroup = "";//店铺税组 允许为空
                //salestotal.vipcode = ""; //VIP 卡号 允许为空
                //salestotal.status = "";//状态 允许为空 保留
                //10:新增/ 20:付款/30:付款取消
                ///40:订单取消
                //在【销售开单查询】中用于显示
                //salestotal.totaldiscount = "";//整单折扣信息 允许为空
                //salestotal.ttltaxamount1 = "";//总税额1   允许为空
                //salestotal.ttltaxamount2 = "";//总税额2   允许为空
                //salestotal.originalamount = "";//原始金额 允许为空 decimal { 4 }  
                //salestotal.priceincludetax = "";//售价是否含税 允许为空
                //salestotal.reservedocno = "";//销售预留库存单号 允许为空
                //salestotal.salesman = "";//销售员 允许为空
                //salestotal.orgtxdocno = "";//原销售单号  允许为空 长度:30
                ////专柜原销售单号
                ////如果提供了【原销售单号】，Web
                ////服务系统判断此单号【是否存
                ////在】或者【是否已经退货】，如
                ////果【不存在】或者【已退货】，
                ////Web 服务系统返回错误信息
                //salestotal.localstorecode = "";//本地店铺号 允许为空 
                //salestotal.orgtxdate_yyyymmdd = "";//原交易日期 允许为空 长度:8
                //固定格式：YYYYMMDD
                //退货时，原交易日期
                //如果是按单退货，提供此日期
                #endregion
                salestotal.changeamount = 0;//找零金额 允许为空 需给零值
                salestotal.cashier = order.OptNo1;//收银员编号 否
                salestotal.issueby = order.PortNo1;//创建人 否
                salestotal.issuedate_yyyymmdd = DateTime.Parse(order.OptTime1).ToString("yyyyMMdd");//"20170408";//创建日期 否 长度:8
                salestotal.txdate_yyyymmdd = DateTime.Parse(order.OptTime1).ToString("yyyyMMdd");// 交易日期 固定格式：YYYYMMDD 否
                salestotal.issuetime_hhmmss = DateTime.Parse(order.OptTime1).ToString("HHmmss");// "140601";//创建时间 否 长度:6 
                salestotal.txtime_hhmmss = DateTime.Parse(order.OptTime1).ToString("HHmmss");// "140601";//交易时间 否 固定格式：HHMMSS
                salestotal.mallid = mallid;//商场编号 长度:4 商场提供固定值 否
                salestotal.mallitemcode = itemcode;//RMS 货号 否  长度:30  Web 服务系统校验货号是否有效
                salestotal.netamount = order.TtlPrice1;// 销售净金额 decimal { 4 }  否
                salestotal.netqty = itemList.Where(n => n.ExchNo == order.ExchNo1).ToList().Count();//净数量 否 销售总数量 decimal { 4 }
                salestotal.paidamount = order.CashPayed1;//付款金额 否
                salestotal.salestype = "SA";// 单据类型 否 SA:店内销售 SR:店内退货/取消交易
                //长度:2
                //销售总金额为正数时，销售类型
                //为 SA；销售总金额为负数时，销
                //单类型为 SR；销售总结 SZ 单据
                //类型是为每日销售总结时发送
                //核对整日销售交易总金额的
                salestotal.sellingamount = order.TtlPrice1;//销售金额 否 decimal { 4 } 
                salestotal.storecode = storeCode;//店铺号 否
                salestotal.tillid = "01";//收银机号 可用 01 或者 02 表示
                //如果专柜只有一台收银机就用
                //01 表示，如果有两台则第二台用
                //02 表示，依次类推
                //Web 服务系统需要校验该收银机
                //编号的有效性
                salestotal.ttpossalesdocno = "";//
                salestotal.txdocno = order.ExchNo1;//销售单号 否 长度:30 专柜销售单号

                request.salestotal = salestotal;
                #endregion

                int qtys = 0;
                #region 销售开单明细表
                //销售开单明细表
                // request.salesitems;
                List<OrderItem> items = itemList.Where(n => n.ExchNo == order.ExchNo1).ToList();
                int leng = items.Count();
                salesitem[] salesitems = new salesitem[leng];
                for (int i = 0; i < items.Count; i++)
                {
                    OrderItem it = items[i];
                    salesitem item = new salesitem();
                    it.ItemNo =itemcode;
                    #region 允许为空
                    //item.bonusearn = 0;// 获得积分 允许为空
                    //item.colorcode = "";// 商品颜色 允许为空
                    //item.coupongroup = "";//优惠券组 允许为空
                    //item.couponnumber = "";//优惠劵号码 允许为空
                    //item.coupontype = "";//优惠劵类型 允许为空
                    //item.exstk2sales = 1;//库存销售比例 允许为空 需给 1 值
                    //item.extendparam = "";//扩展参数 允许为空
                    //item.invttype = "1";//库存类型 允许为空 0:坏货退回/1:好货退回
                    //默认为 1
                    //主要用于店内退货，在 PDA 店内
                    //退货时选择库存类型，单品后台
                    //系统根据库存类型进行控制是
                    //否增加库存
                    //item.isdeposit = "";//是否定金单  允许为空
                    //item.iswholesale = "";//是否批发 允许为空
                    //item.priceapprove = "";//允许改价 允许为空
                    //item.pricemode = "";//价格模式  允许为空
                    //item.promotion = "";//商品促销信息 允许为空
                    //item.tax = "";//商品税信息 允许为空
                    //item.sellingprice = 0;//售价 允许为空 需给零值
                    //item.serialnum = "";// 序列号 允许为空
                    //item.sizecode = "";//商品尺码 允许为空
                    //item.totaldiscountless = 0;//整单折扣差额 需给零值 允许为空
                    //item.totaldiscountless1 = 0;//整单折扣差额 1 需给零值 允许为空
                    //item.totaldiscountless2 = 0;//整单折扣差额 2 需给零值 允许为空
                    //item.vipdiscountless = 0;//VIP 折扣差额  需给零值 允许为空
                    //item.vipdiscountpercent = 0;//VIP 折扣率 需给零值    允许为空
                    //item.refundreasoncode = "";//退货原因  允许为空
                    //item.salesitemremark = "";// 交易明细备注  允许为空
                    //item.itemdiscount = "";//单品折扣信息 允许为空
                    //item.itemlotnum = "";// 商品批次 允许为空
                    //item.originalprice = 0;//原始售价  允许为空 需给零值
                    #endregion
                    item.iscounteritemcode = "1";//是否专柜货号 否 默认为 1 长度:1
                    item.lineno = int.Parse(order.ExchID1);//行号 否
                    item.mallitemcode = it.ItemNo;//货号 否
                    item.plucode = it.ItemNo;//商品内部编号 否 同 itemcode
                    item.counteritemcode = it.ItemNo;// 专柜货号 否 如果不区分【专柜货号】与【商品编号】二者赋予相同值
                    item.itemcode = it.ItemNo;//商品编号 否 Web 服务系统校验货号是否有效同 mallitemcode ，counteritemcode
                    item.netamount = it.Price;// 销售净金额 decimal { 4 }  否
                    item.qty = it.Qty;//数量 否 decimal { 4 }  
                    item.storecode = storeCode;//店铺号 否 
                    salesitems[i] = item;
                    qtys += it.Qty;  
                }
                sumCount += qtys;
                #endregion
                request.salesitems = salesitems;
                #region 销售开单付款明细表
                //销售开单付款明细表
                //request.salestenders;
                WebReference.salestender salestender = new WebReference.salestender();
                salestender.baseamount = order.TtlPrice1;//本位币金额  否 同 payamount
                //salestender.excessamount = ""; //超额金额 可以为空
                //salestender.extendparam = "";//扩展参数 可以为空
                //salestender.tendertype = "";//付款类型 可以为空
                salestender.lineno = int.Parse(order.ExchID1);//行号 否
                salestender.payamount = order.TtlPrice1;//付款金额 否
                //salestender.remark = "";//备注 可以为空
                //salestender.tendercategory = "";//付款种类 可以为空
                salestender.tendercode = "CH";//付款代码 否 长度:2
                //CH----现金
                //CI----国内银行卡
                //CO----国外银行卡
                //OT-----其他付款方式。
                //接口数据应在 TenderCode 付款
                //方式中填写对应方式付款实际
                //金额，无对应付款方式时在其他
                //付款方式字段填写剩余付款方
                //式金额的合计；若交易类型为 SZ 
                //可以默认为 OT
                //Web 服务系统需要校验付款方式
                //编号有效性
                request.salestenders = new WebReference.salestender[] { salestender };
                #endregion

                #region 销售开单配送表
                //销售开单配送表 允许为空
                // request.salesdlvy
                request.salesdlvy = new salesdelivery();
                #endregion

                //返回码（<responsecode>short</responsecode>）为【0】，表示调用 Web Service 成功。
                //交易被完整接纳。
                //返回码（<responsecode>short</responsecode>）为非【0】，表示调用 Web Service 不
                //成功。软件开发商收到此返回信息清除重传交易队列销售资料。
                //其他返回码表示不成功，不成功信息从（ <responsemessage>string</responsemessage> ）获取。
                //软件开发商收到其他返回码请将未成功传送的交易放入重传交易队列。

                //string xml = Serializer(typeof(postsalescreaterequest), request);

                WebReference.postsalescreateresponse response = sales.postsalescreate(request);
                short resposeCode = response.header.responsecode;
                string responseMessage = response.header.responsemessage;
                if (resposeCode == 0)
                {
                    //调用成功
                    responseMessage = "同步成功";
                }
                else
                {
                    responseMessage = "错误码：" + resposeCode + ";错误信息：" + responseMessage;  
                    //调用失败
                }
                k++;
                this.Dispatcher.BeginInvoke(new Action(
                 () =>
                 {
                     btnDataSync.Content = "正在同步数据第(" + (k) + "/" + (orderList.Count + orderRejectList.Count) + ")";
                     this.Body.Text = "同步销售总金额：" + String.Format("{0:N2}",sumPrice) + "；总数量：" + sumCount;
              
                     logs.Items.Add(new { NID = k, ExchNo1 = order.ExchNo1, TtlPrice1 = String.Format("{0:N2}", order.TtlPrice1), qtys = qtys, responseMessage = responseMessage });
                 }));
            }
            //this.Dispatcher.BeginInvoke(new Action(
            // () =>
            // {
            //     this.btnDataSync.Content = "开始同步";
            //     MessageBox.Show("同步总金额：" + String.Format("{0:N2}", sumPrice) + "；总数量：" + sumCount); 
            // }));

            //退货

            SendRejectData(orderList, itemList, orderRejectList, orderRejectItemList);

            }
            catch (Exception ex)
            {
                MessageBox.Show("同步异常：" +ex.Message); 
            }
        }
        /// <summary>
        /// 提交退款的记录
        /// </summary>
        /// <param name="orderList"></param>
        /// <param name="itemList"></param>
        public void SendRejectData(List<Order> orderList, List<OrderItem> itemList,List<Order> orderRejectList,List<OrderItem> itemRejectList)
        {

            int k = 0;
            decimal sumPrice = 0;
            int sumCount = 0;
            foreach (Order order in orderRejectList)
            {
                if(order.Remark1 == ""){
                    continue;
                }
                string orderIDOLD = order.Remark1.Split('流')[0]; ;//原单号: S001170330171356  流水号: 164;
                orderIDOLD = orderIDOLD.Split(':')[1];
                orderIDOLD = orderIDOLD.Trim();

                Order oldOrder = GetRejectOrder(orderIDOLD);
                if(oldOrder == null)
                {
                    continue;
                }

                WebReference.sales sales = new WebReference.sales();

                WebReference.postsalescreaterequest request = new WebReference.postsalescreaterequest();

                #region header;//标头
                requestheader header = new requestheader();
                header.lang = ""; //语言
                header.pageno = 0;//页数 
                header.pagerecords = 0;//每页记录数
                header.updatecount = 0;//每次更新记录数 

                header.licensekey = licensekey;
                header.password = password;
                header.username = username;
                header.messageid = "332";//固定值：332
                header.messagetype = "SALESDATA";//固定值：SALESDATA
                header.version = "V332M";//固定值：
                request.header = header;
                #endregion

                sumPrice += order.TtlPrice1;

                #region 销售开单主表
                saleshdr salestotal = new saleshdr();
                // 销售开单主表 
                #region 可以为空
                //salestotal.buyerremark = "";//卖家备注 允许为空
                //salestotal.coupongroup = "";//优惠券组 允许为空
                //salestotal.couponnumber = "";//优惠券号码 允许为空
                //salestotal.couponqty = "";//优惠券数量 允许为空
                //salestotal.coupontype = "";//优惠券类型 允许为空
                //salestotal.demographiccode = "";//顾客统计代码 允许为空
                //salestotal.demographicdata = "";//顾客统计值 允许为空
                //salestotal.ecorderno = "";//网购订单号 允许为空
                //salestotal.extendparam = "";//扩展参数 允许为空
                //salestotal.invoicecontent = "";//发票内容 允许为空
                //salestotal.invoicetitle = "";//发票抬头 允许为空
                //salestotal.orderremark = "";//交易备注 允许为空
                //salestotal.shoptaxgroup = "";//店铺税组 允许为空
                //salestotal.vipcode = ""; //VIP 卡号 允许为空
                //salestotal.status = "";//状态 允许为空 保留
                //10:新增/ 20:付款/30:付款取消
                ///40:订单取消
                //在【销售开单查询】中用于显示
                //salestotal.totaldiscount = "";//整单折扣信息 允许为空
                //salestotal.ttltaxamount1 = "";//总税额1   允许为空
                //salestotal.ttltaxamount2 = "";//总税额2   允许为空
                //salestotal.priceincludetax = "";//售价是否含税 允许为空
                //salestotal.reservedocno = "";//销售预留库存单号 允许为空
                //salestotal.salesman = "";//销售员 允许为空
                salestotal.orgtxdocno = orderIDOLD;//原销售单号  允许为空 长度:30
                salestotal.orgstorecode = storeCode;//原交易店铺号 允许为空 退货时，原交易店铺号 如果是按单退货，提供此店铺号
                salestotal.orgtillid = "01";// 原收银机号 允许为空 退货时，原收银机号  如果是按单退货，提供此收银机号
                //专柜原销售单号
                //如果提供了【原销售单号】，Web
                //服务系统判断此单号【是否存
                //在】或者【是否已经退货】，如
                //果【不存在】或者【已退货】，
                //Web 服务系统返回错误信息
                //salestotal.localstorecode = "";//本地店铺号 允许为空 
                salestotal.originalamount = oldOrder.TtlPrice1;//原始金额 允许为空 decimal { 4 }  
                salestotal.orgtxdate_yyyymmdd = DateTime.Parse(oldOrder.OptTime1).ToString("yyyyMMdd");//原交易日期 允许为空 长度:8
                //固定格式：YYYYMMDD
                //退货时，原交易日期
                //如果是按单退货，提供此日期
                #endregion
                salestotal.changeamount = 0;//找零金额 允许为空 需给零值
                salestotal.cashier = order.OptNo1;//收银员编号 否
                salestotal.issueby = order.PortNo1;//创建人 否
                salestotal.issuedate_yyyymmdd = DateTime.Parse(order.OptTime1).ToString("yyyyMMdd");//"20170408";//创建日期 否 长度:8
                salestotal.txdate_yyyymmdd = DateTime.Parse(order.OptTime1).ToString("yyyyMMdd");// 交易日期 固定格式：YYYYMMDD 否
                salestotal.issuetime_hhmmss = DateTime.Parse(order.OptTime1).ToString("HHmmss");// "140601";//创建时间 否 长度:6 
                salestotal.txtime_hhmmss = DateTime.Parse(order.OptTime1).ToString("HHmmss");// "140601";//交易时间 否 固定格式：HHMMSS
                salestotal.mallid = mallid;//商场编号 长度:4 商场提供固定值 否
                salestotal.mallitemcode = "";//RMS 货号 否  长度:30  Web 服务系统校验货号是否有效
                salestotal.netamount = -order.TtlPrice1;// 销售净金额 decimal { 4 }  否
                salestotal.netqty = itemList.Where(n => n.ExchNo == order.ExchNo1).ToList().Count();//净数量 否 销售总数量 decimal { 4 }
                salestotal.paidamount = -order.CashPayed1;//付款金额 否
                salestotal.salestype = "SR";// 单据类型 否 SA:店内销售 SR:店内退货/取消交易
                //长度:2
                //销售总金额为正数时，销售类型
                //为 SA；销售总金额为负数时，销
                //单类型为 SR；销售总结 SZ 单据
                //类型是为每日销售总结时发送
                //核对整日销售交易总金额的
                salestotal.sellingamount = -order.TtlPrice1;//销售金额 否 decimal { 4 } 
                salestotal.storecode = storeCode;//店铺号 否
                salestotal.tillid = "01";//收银机号 可用 01 或者 02 表示
                //如果专柜只有一台收银机就用
                //01 表示，如果有两台则第二台用
                //02 表示，依次类推
                //Web 服务系统需要校验该收银机
                //编号的有效性
                salestotal.ttpossalesdocno = "";//
                salestotal.txdocno = order.ExchNo1;//销售单号 否 长度:30 专柜销售单号

                request.salestotal = salestotal;
                #endregion

                int qtys = 0;
                #region 销售开单明细表
                //销售开单明细表
                // request.salesitems;
                List<OrderItem> items = itemRejectList.Where(n => n.ExchNo == order.ExchNo1).ToList();
                int leng = items.Count();
                salesitem[] salesitems = new salesitem[leng];
                for (int i = 0; i < items.Count; i++)
                {
                    OrderItem it = items[i];
                    salesitem item = new salesitem();
                    it.ItemNo = itemcode;// "A000011";
                    #region 允许为空
                    //item.bonusearn = 0;// 获得积分 允许为空
                    //item.colorcode = "";// 商品颜色 允许为空
                    //item.coupongroup = "";//优惠券组 允许为空
                    //item.couponnumber = "";//优惠劵号码 允许为空
                    //item.coupontype = "";//优惠劵类型 允许为空
                    //item.exstk2sales = 1;//库存销售比例 允许为空 需给 1 值
                    //item.extendparam = "";//扩展参数 允许为空
                    //item.invttype = "1";//库存类型 允许为空 0:坏货退回/1:好货退回
                    //默认为 1
                    //主要用于店内退货，在 PDA 店内
                    //退货时选择库存类型，单品后台
                    //系统根据库存类型进行控制是
                    //否增加库存
                    //item.isdeposit = "";//是否定金单  允许为空
                    //item.iswholesale = "";//是否批发 允许为空
                    //item.priceapprove = "";//允许改价 允许为空
                    //item.pricemode = "";//价格模式  允许为空
                    //item.promotion = "";//商品促销信息 允许为空
                    //item.tax = "";//商品税信息 允许为空
                    //item.sellingprice = 0;//售价 允许为空 需给零值
                    //item.serialnum = "";// 序列号 允许为空
                    //item.sizecode = "";//商品尺码 允许为空
                    //item.totaldiscountless = 0;//整单折扣差额 需给零值 允许为空
                    //item.totaldiscountless1 = 0;//整单折扣差额 1 需给零值 允许为空
                    //item.totaldiscountless2 = 0;//整单折扣差额 2 需给零值 允许为空
                    //item.vipdiscountless = 0;//VIP 折扣差额  需给零值 允许为空
                    //item.vipdiscountpercent = 0;//VIP 折扣率 需给零值    允许为空
                    //item.refundreasoncode = "";//退货原因  允许为空
                    //item.salesitemremark = "";// 交易明细备注  允许为空
                    //item.itemdiscount = "";//单品折扣信息 允许为空
                    //item.itemlotnum = "";// 商品批次 允许为空
                    //item.originalprice = 0;//原始售价  允许为空 需给零值
                    #endregion
                    item.iscounteritemcode = "1";//是否专柜货号 否 默认为 1 长度:1
                    item.lineno = int.Parse(order.ExchID1);//行号 否
                    item.mallitemcode = it.ItemNo;//货号 否
                    item.plucode = it.ItemNo;//商品内部编号 否 同 itemcode
                    item.counteritemcode = it.ItemNo;// 专柜货号 否 如果不区分【专柜货号】与【商品编号】二者赋予相同值
                    item.itemcode = it.ItemNo;//商品编号 否 Web 服务系统校验货号是否有效同 mallitemcode ，counteritemcode
                    item.netamount = -it.Price;// 销售净金额 decimal { 4 }  否
                    item.qty = it.Qty;//数量 否 decimal { 4 }  
                    item.storecode = storeCode;//店铺号 否 
                    salesitems[i] = item;
                    qtys += it.Qty;
                }
                sumCount += qtys;
                #endregion
                request.salesitems = salesitems;
                #region 销售开单付款明细表
                //销售开单付款明细表
                //request.salestenders;
                WebReference.salestender salestender = new WebReference.salestender();
                salestender.baseamount = -order.TtlPrice1;//本位币金额  否 同 payamount
                //salestender.excessamount = ""; //超额金额 可以为空
                //salestender.extendparam = "";//扩展参数 可以为空
                //salestender.tendertype = "";//付款类型 可以为空
                salestender.lineno = int.Parse(order.ExchID1);//行号 否
                salestender.payamount = -order.TtlPrice1;//付款金额 否
                //salestender.remark = "";//备注 可以为空
                //salestender.tendercategory = "";//付款种类 可以为空
                salestender.tendercode = "CH";//付款代码 否 长度:2
                //CH----现金
                //CI----国内银行卡
                //CO----国外银行卡
                //OT-----其他付款方式。
                //接口数据应在 TenderCode 付款
                //方式中填写对应方式付款实际
                //金额，无对应付款方式时在其他
                //付款方式字段填写剩余付款方
                //式金额的合计；若交易类型为 SZ 
                //可以默认为 OT
                //Web 服务系统需要校验付款方式
                //编号有效性
                request.salestenders = new WebReference.salestender[] { salestender };
                #endregion

                #region 销售开单配送表
                //销售开单配送表 允许为空
                // request.salesdlvy
                request.salesdlvy = new salesdelivery();
                #endregion

                //返回码（<responsecode>short</responsecode>）为【0】，表示调用 Web Service 成功。
                //交易被完整接纳。
                //返回码（<responsecode>short</responsecode>）为非【0】，表示调用 Web Service 不
                //成功。软件开发商收到此返回信息清除重传交易队列销售资料。
                //其他返回码表示不成功，不成功信息从（ <responsemessage>string</responsemessage> ）获取。
                //软件开发商收到其他返回码请将未成功传送的交易放入重传交易队列。

                WebReference.postsalescreateresponse response = sales.postsalescreate(request);
                short resposeCode = response.header.responsecode;
                string responseMessage = response.header.responsemessage;
                if (resposeCode == 0)
                {
                    //调用成功
                    responseMessage = "同步成功";
                }
                else
                {
                    responseMessage = "错误码：" + resposeCode + ";错误信息：" + responseMessage;  
                    //调用失败
                }
                k++;
                this.Dispatcher.BeginInvoke(new Action(
                 () =>
                 {
                     btnDataSync.Content = "正在同步数据第(" + (k) + "/" + (orderList.Count + orderRejectList.Count) + ")";
                     this.RejectText.Text = "同步退款总金额：" + String.Format("{0:N2}", -sumPrice) + "；总数量：" + sumCount;

                     logs.Items.Add(new { NID = k, ExchNo1 = order.ExchNo1 + "(原始单号：" + orderIDOLD + ")", TtlPrice1 = String.Format("{0:N2}", -order.TtlPrice1), qtys = qtys, responseMessage = responseMessage });
                 }));
            }
            this.Dispatcher.BeginInvoke(new Action(
             () =>
             {
                 this.btnDataSync.Content = "开始同步";
                 btnDataSync.IsEnabled = true;
                 MessageBox.Show("同步完成");
                 //MessageBox.Show("同步总金额：" + String.Format("{0:N2}", sumPrice) + "；总数量：" + sumCount);
             }));
        }
        [DllImport("Iphlpapi.dll")]
        static extern int SendARP(Int32 DestIP, Int32 SrcIP, ref Int64 MacAddr, ref Int32 PhyAddrLen);
        /// <summary>  
        /// SendArp获取MAC地址  
        /// </summary>  
        /// <returns></returns>  
        public string GetMacAddressBySendARP()
        {
            StringBuilder strReturn = new StringBuilder();
            try
            {
                System.Net.IPHostEntry Tempaddr = (System.Net.IPHostEntry)Dns.GetHostByName(Dns.GetHostName());
                System.Net.IPAddress[] TempAd = Tempaddr.AddressList;
                Int32 remote = (int)TempAd[0].Address;
                Int64 macinfo = new Int64();
                Int32 length = 6;
                SendARP(remote, 0, ref macinfo, ref length);

                string temp = System.Convert.ToString(macinfo, 16).PadLeft(12, '0').ToUpper();

                int x = 12;
                for (int i = 0; i < 6; i++)
                {
                    if (i == 5) { strReturn.Append(temp.Substring(x - 2, 2)); }
                    else { strReturn.Append(temp.Substring(x - 2, 2) + "-"); }
                    x -= 2;
                }

                return strReturn.ToString();
            }
            catch
            {
                return "";
            }
        } 
        /// <summary>
        /// 时间刷新
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        void timer_Tick(object sender, EventArgs e)
        {
            string time = string.Concat("  时间：  ", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
            this.CurrentDateText.Text = string.Concat(time, ip);
        }

        /// <summary>
        /// 开始同步数据
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        public void StartSynData(object sender, RoutedEventArgs e)
        {
            btnDataSync.Content = "正在同步数据...";
            btnDataSync.IsEnabled = false;
            StartTimeText.Text = "同步开始时间：" + DateTime.Now;
            orderList = GetOrderList();
            orderItemList = GetOrderItemList();
            orderRejectList = GetRejectOrderList();
            orderRejectItemList = GetRejectOrderItemList();
              if (orderList == null)
              {
                  this.btnDataSync.Content = "开始同步";
                MessageBox.Show("数据为空，不需要同步。");
            }
            else
            {
            Thread thread = new Thread(Start);
            thread.Start();
            }
        }

        public void Start()
        {
          
                SendData(orderList, orderItemList); 
           
        } 

        /// <summary>
        /// 获取本地IP
        /// </summary>
        /// <returns></returns>
        protected string GetIP()  
        {
            IPHostEntry ipHost = Dns.Resolve(Dns.GetHostName());
            IPAddress ipAddr = ipHost.AddressList[0];
            return ipAddr.ToString();
        }
        #region 序列化
        /// <summary>
        /// 序列化
        /// </summary>
        /// <param name="type">类型</param>
        /// <param name="obj">对象</param>
        /// <returns></returns>
        public static string Serializer(Type type, object obj)
        {
            MemoryStream Stream = new MemoryStream();
            XmlSerializer xml = new XmlSerializer(type);
            try
            {
                //序列化对象
                xml.Serialize(Stream, obj);
            }
            catch (InvalidOperationException)
            {
                throw;
            }
            Stream.Position = 0;
            StreamReader sr = new StreamReader(Stream);
            string str = sr.ReadToEnd();

            sr.Dispose();
            Stream.Dispose();

            return str;
        }

        #endregion
    }
    public class Order
    {
        private string ExchID;

        public string ExchID1
        {
            get { return ExchID; }
            set { ExchID = value; }
        }
        private string ExchNo;

        public string ExchNo1
        {
            get { return ExchNo; }
            set { ExchNo = value; }
        }
        private string HelpNo;

        public string HelpNo1
        {
            get { return HelpNo; }
            set { HelpNo = value; }
        }
        private decimal TtlPrice;

        public decimal TtlPrice1
        {
            get { return TtlPrice; }
            set { TtlPrice = value; }
        }
        private string MnlDisc;

        public string MnlDisc1
        {
            get { return MnlDisc; }
            set { MnlDisc = value; }
        }
        private string DiscRate;

        public string DiscRate1
        {
            get { return DiscRate; }
            set { DiscRate = value; }
        }
        private string TtlFavor;

        public string TtlFavor1
        {
            get { return TtlFavor; }
            set { TtlFavor = value; }
        }

        private string MnlTip;

        public string MnlTip1
        {
            get { return MnlTip; }
            set { MnlTip = value; }
        }
        public string TipAmt;
        private string TktPayed;

        public string TktPayed1
        {
            get { return TktPayed; }
            set { TktPayed = value; }
        }
        private string CardPayed;

        public string CardPayed1
        {
            get { return CardPayed; }
            set { CardPayed = value; }
        }
        private string ChqPayed;

        public string ChqPayed1
        {
            get { return ChqPayed; }
            set { ChqPayed = value; }
        }
        private string BankPayed;

        public string BankPayed1
        {
            get { return BankPayed; }
            set { BankPayed = value; }
        }
        private decimal CashPayed;

        public decimal CashPayed1
        {
            get { return CashPayed; }
            set { CashPayed = value; }
        }


        private string BillStatus;

        public string BillStatus1
        {
            get { return BillStatus; }
            set { BillStatus = value; }
        }
        private string SvcPsn;

        public string SvcPsn1
        {
            get { return SvcPsn; }
            set { SvcPsn = value; }
        }
        private string OptTime;

        public string OptTime1
        {
            get { return OptTime; }
            set { OptTime = value; }
        }
        private string StlID;

        public string StlID1
        {
            get { return StlID; }
            set { StlID = value; }
        }
        private string OptNo;

        public string OptNo1
        {
            get { return OptNo; }
            set { OptNo = value; }
        }
        private string PortNo;

        public string PortNo1
        {
            get { return PortNo; }
            set { PortNo = value; }
        }
        public string TrunID;


        private string Remark;

        public string Remark1
        {
            get { return Remark; }
            set { Remark = value; }
        }
        private string TranStatus;

        public string TranStatus1
        {
            get { return TranStatus; }
            set { TranStatus = value; }
        }
    }
    public class OrderItem
    {
        private string exchNo;

        public string ExchNo
        {
            get { return exchNo; }
            set { exchNo = value; }
        }

        private string itemNo;

        public string ItemNo
        {
            get { return itemNo; }
            set { itemNo = value; }
        }
        private decimal realPrice;

        public decimal RealPrice
        {
            get { return realPrice; }
            set { realPrice = value; }
        }

        private decimal price;

        public decimal Price
        {
            get { return price; }
            set { price = value; }
        }
        private int qty;

        public int Qty
        {
            get { return qty; }
            set { qty = value; }
        }
    }
}
