using System;
using System.Runtime.InteropServices;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Preactor;
using Preactor.Interop.PreactorObject;

namespace ApsPbxLibrary
{
    [Guid("FAA5E07F-E3FC-450E-9669-13BE7A54059E")]
    [ComVisible(true)]
    public interface IPbxAction
    {
        int Run(ref PreactorObj preactorComObject, ref object pespComObject);
    }

    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    [Guid("42B62270-811F-4311-AA2E-ECBBFFA27F39")]
    public class PbxAction : IPbxAction
    {
        const string BATCH_SIZE_ATTR = "Numerical Attribute 1";
        const string DEMAND_ORDERNO_ATTR = "String Attribute 1";

        /// <summary>
        /// Счетчик номеров заказа
        /// </summary>
        int orderNoIndex = 1;

        /// <summary>
        /// Формирование всех ПЗ с учетом партийности
        /// </summary>
        /// <remarks>
        /// Допущения:
        /// 1. Есть PESP "SMC Default", запускающий связку по Default
        /// 2. Наследуются Priority, Due Date и обозначение Demand в Orders.String Attribute 1
        /// 3. Products.Numerical Attribute 1 - размер партии для деления ПЗ
        /// </remarks>
        public int Run(ref PreactorObj preactorComObject, ref object pespComObject)
        {
            CU.Preactor = PreactorFactory.CreatePreactorObject(preactorComObject);
            CU.Core = EventScriptsFactory.CreateEventScriptCoreObject(preactorComObject, pespComObject);
            try
            {
                int level = 1;
                int ordersCnt;
                int newOrderCount = 0;
                do
                {
                    ordersCnt = newOrderCount;
                    CU.RunSmc(CU.SmcRule.Default);
                    CU.Preactor.DisplayStatus("Формирование ПЗ", $"уровень {level++}");
                    PbxStep();
                    CU.Preactor.Redraw();
                    newOrderCount = CU.Preactor.RecordCount("Orders");
                } while (ordersCnt != newOrderCount);
                CU.Preactor.DestroyStatus();
            }
            catch (AbortException)
            {
                MessageBox.Show("Выполнение прервано", "Сообщение", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            return 0;
        }

        /// <summary>
        /// Развертывание одного уровня дефицитов
        /// </summary>
        void PbxStep()
        {
            int shortagesCnt = CU.Preactor.RecordCount("Shortages");
            for (int shortageRecNo = 1; shortageRecNo <= shortagesCnt; shortageRecNo++)
            {
                CU.Preactor.UpdateStatus(shortageRecNo, shortagesCnt);
                try
                {
                    Shortage shortage = Shortage.ByRecNo(shortageRecNo);
                    if (!shortage.IsIgnored())
                    {
                        MakeOrdersForShortage(shortage);
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Ошибка при обработке дефицита ID = {CU.Preactor.ReadFieldInt("Shortages", "Number", shortageRecNo)}:\n{ex.Message}",
                        "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
        }

        /// <summary>
        /// Получение очередного незанятого номера заказа
        /// </summary>
        /// <returns></returns>
        string GetNewOrderNo()
        {
            string res;
            do
            {
                res = $"ПЗ-{orderNoIndex++}";
            } while (CU.Preactor.FindMatchingRecord("Orders", "Order No.", 0, res) > 0);
            return res;
        }

        /// <summary>
        /// Формирование производственных заказов для восполнения дефицита
        /// </summary>
        /// <param name="shortage">дефицит</param>
        /// <returns>
        /// true - заказы созданы
        /// false - заказы не созданы
        /// </returns>
        bool MakeOrdersForShortage(Shortage shortage)
        {
            int productRecNo = CU.Preactor.FindMatchingRecord("Products", "Part No.", 0, shortage.PartNo);
            if (productRecNo <= 0)
                return false;
            double batchSize = GetBatchSize(productRecNo) ?? shortage.Quantity;
            while (shortage.Quantity > 0)
            {
                if (batchSize > shortage.Quantity)
                    batchSize = shortage.Quantity;
                CreateWorkOrder(shortage, shortage.PartNo, batchSize);
                shortage.Quantity -= batchSize;
            }
            return true;
        }

        /// <summary>
        /// Получение размера партии для заданной номенклатуры
        /// </summary>
        /// <param name="partNo">номер записи в Products</param>
        /// <returns>Размер партии или null, если он не указан</returns>
        double? GetBatchSize(int productRecNo)
        {
            double res = CU.Preactor.ReadFieldDouble("Products", BATCH_SIZE_ATTR, productRecNo);
            return res > 0 ? (double?)res : null;
        }

        /// <summary>
        /// Создание производственного заказа на основе дефицита по заданному ТП в заданном количестве
        /// </summary>
        /// <param name="shortage">данные о восполняемом дефиците</param>
        /// <param name="routeCode">применяемый ТП</param>
        /// <param name="quantity">размер создаваемого заказа</param>
        void CreateWorkOrder(Shortage shortage, string routeCode, double quantity)
        {
            string orderNo = GetNewOrderNo();
            int orderRecNo = CU.Preactor.CreateRecord("Orders");
            CU.Preactor.WriteField("Orders", "Order No.", orderRecNo, orderNo);
            CU.Preactor.WriteField("Orders", "Part No.", orderRecNo, routeCode);
            CU.Preactor.WriteField("Orders", "Quantity", orderRecNo, quantity);
            CU.Preactor.WriteField("Orders", "Due Date", orderRecNo, shortage.DueDate);
            CU.Preactor.ExpandJob("Orders", orderRecNo);
            CU.Preactor.WriteField("Orders", DEMAND_ORDERNO_ATTR, orderRecNo, shortage.DemandOrderNo);
            CU.Preactor.WriteField("Orders", "Priority", orderRecNo, shortage.Priority);
        }

        class Shortage
        {
            public int ExternalDemandOrderId { get; private set; }
            public int InternalDemandOrderId { get; private set; }
            public string PartNo { get; private set; }
            public double Quantity { get; set; }
            public string DemandOrderNo { get; private set; }
            public DateTime DueDate { get; private set; }
            public int Priority { get; private set; }

            public static Shortage ByRecNo(int recNo)
            {
                Shortage res = new Shortage()
                {
                    ExternalDemandOrderId = CU.Preactor.ReadFieldInt("Shortages", "External Demand Order", recNo),
                    InternalDemandOrderId = CU.Preactor.ReadFieldInt("Shortages", "Internal Demand Order", recNo),
                    PartNo = CU.Preactor.ReadFieldString("Shortages", "Part No.", recNo),
                    Quantity = CU.Preactor.ReadFieldDouble("Shortages", "Shortage Quantity", recNo)
                };
                if (res.ExternalDemandOrderId > 0)
                {
                    PrimaryKey demandKey = new PrimaryKey(res.ExternalDemandOrderId);
                    res.DemandOrderNo = CU.Preactor.ReadFieldString("Demand", "Order No.", demandKey);
                    res.DueDate = CU.Preactor.ReadFieldDateTime("Demand", "Demand Date", demandKey);
                    res.Priority = CU.Preactor.ReadFieldInt("Demand", "Priority", demandKey);
                }
                else
                {
                    PrimaryKey orderKey = new PrimaryKey(res.InternalDemandOrderId);
                    res.DemandOrderNo = CU.Preactor.ReadFieldString("Orders", DEMAND_ORDERNO_ATTR, orderKey);
                    res.DueDate = CU.Preactor.ReadFieldDateTime("Orders", "Due Date", orderKey);
                    res.Priority = CU.Preactor.ReadFieldInt("Orders", "Priority", orderKey);
                }
                return res;
            }

            public bool IsIgnored()
            {
                int ignoreRecNo = CU.Preactor.FindMatchingRecord("Ignore Shortages", 0,
                    $"(({{#External Demand Order}} == {ExternalDemandOrderId}) && ({{#Internal Demand Order}} == {InternalDemandOrderId}) && " +
                    $"(~{{$Part No.}}~ == ~{PartNo}~) && ({{#Ignore Shortages}} == 1))");
                return ignoreRecNo > 0;
            }

            public override string ToString()
            {
                if (ExternalDemandOrderId > 0)
                    return $"DemandId = {ExternalDemandOrderId}, Part No. = '{PartNo}', Quantity = {Quantity}";
                else
                    return $"WorkOrderId = {InternalDemandOrderId}, Part No. = '{PartNo}', Quantity = {Quantity}";
            }
        }
    }
}
