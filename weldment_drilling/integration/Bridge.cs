using System;
using System.Runtime.InteropServices;
using AutoSizeDrilling;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace Macro.DrillingDimensionsIntegration
{
    [ComVisible(true)]
    [Guid("31b8075c-5011-4d82-9e3f-d3c6da09be2c")]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface IBridge
    {
        [DispId(1)]
        string BuildId { get; }

        [DispId(2)]
        [return: MarshalAs(UnmanagedType.Struct)]
        object ProcessSelectedView(
            [MarshalAs(UnmanagedType.IDispatch)] object application,
            [MarshalAs(UnmanagedType.IDispatch)] object expectedView, double targetRight, double targetTop);
    }

    [ComVisible(true)]
    [Guid("93769c05-7e2b-4350-bbbe-36a2a60f52d7")]
    [ProgId("Macro.DrillingDimensions")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IBridge))]
    public sealed class Bridge : IBridge
    {
        public string BuildId => "DRILLING_DIMENSIONS_BRIDGE_R2_20261029";

        public object ProcessSelectedView(object application, object expectedView, double targetRight, double targetTop)
        {
            // The caller supplies its actual SW instance. Do not attach through the ROT:
            // several SOLIDWORKS processes may be running at once.
            var app = application as SldWorks;
            if (app == null)
                throw new ArgumentException("Передайте текущий экземпляр SOLIDWORKS.", nameof(application));
            var expected = expectedView as View;
            if (expected == null)
                throw new ArgumentException("Не передан ожидаемый объект чертёжного вида.", nameof(expectedView));
            var model = app.ActiveDoc as ModelDoc2;
            if (model == null || model.GetType() != (int)swDocumentTypes_e.swDocDRAWING)
                throw new InvalidOperationException("Активный документ должен быть исходным чертежом.");
            var selection = model.SelectionManager as SelectionMgr;
            if (selection == null || selection.GetSelectedObjectCount2(-1) != 1)
                throw new InvalidOperationException("Перед размерами выберите ровно один созданный вид.");
            if (selection.GetSelectedObjectType3(1, -1) != (int)swSelectType_e.swSelDRAWINGVIEWS)
                throw new InvalidOperationException("Выбран объект другого типа вместо чертёжного вида.");
            var selected = selection.GetSelectedObject6(1, -1) as View;
            if (selected == null)
                throw new InvalidOperationException("Не удалось получить выбранный чертёжный вид.");
            if (!ReferenceEquals(selected, expected) &&
                app.IsSame(selected, expected) != (int)swObjectEquality.swObjectSame)
                throw new InvalidOperationException("Выбранный объект не совпал с ожидаемым созданным видом. Имена не используются для сопоставления.");

            // This entry point propagates errors AFTER rollback/mode restoration.
            // The legacy Run(host)/Execute paths can swallow an error and are not used.
            double[] result = new AutoSizeDrill().RunForIntegration(app, targetRight, targetTop);
            ValidateResult(result);
            return result;
        }

        private static void ValidateResult(double[] result)
        {
            // Sheet metres; conservative envelope of dimension anchors plus padding,
            // not a native glyph bounding rectangle. Count is the validated dimensions.
            if (result == null || result.Length != 5)
                throw new InvalidOperationException("AutoSizeDrilling вернул неполный результат обработки вида.");
            foreach (double value in result)
                if (double.IsNaN(value) || double.IsInfinity(value))
                    throw new InvalidOperationException("AutoSizeDrilling вернул неконечные координаты.");
            if (result[0] < 1 || result[0] != Math.Floor(result[0]) ||
                result[1] >= result[3] || result[2] >= result[4])
                throw new InvalidOperationException("AutoSizeDrilling не подтвердил размеры или габариты обработанного вида.");
        }
    }
}
