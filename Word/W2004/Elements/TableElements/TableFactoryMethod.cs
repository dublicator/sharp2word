namespace Word.W2004.Elements.TableElements
{
    /// <summary>
    ///   Here is the logic to decide which instance create and return
    /// </summary>
    public class TableFactoryMethod
    {
        private static TableFactoryMethod _instance;

        private TableFactoryMethod()
        {
        }

        public static TableFactoryMethod Instance
        {
            get { return _instance ?? (_instance = new TableFactoryMethod()); }
        }

        public static ITableItemStrategy GetTableItem(TableEle tableEle)
        {
            if (tableEle == null)
            {
                return null;
            }

            return GetTableEle(tableEle);
        }

        private static ITableItemStrategy GetTableEle(TableEle tableEle)
        {
            if (tableEle.Value.Equals("tableDef"))
            {
                return new TableDefinition();
            }
            else if (tableEle.Value.Equals("th"))
            {
                return new TableHeader();
            }
            else if (tableEle.Value.Equals("td"))
            {
                return new TableCol();
            }
            else
            {
                //if (tableEle.getValue().equals("tf")) {
                return new TableFooter();
            }
        }

    }
}