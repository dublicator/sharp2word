using NUnit.Framework;
using Word.W2004.Elements.TableElements;

namespace Test.W2004
{
    public class TableFactoryMethodTest : Assert
    {
        [Test]
        public void testGetInstance()
        {
            TableFactoryMethod instance = TableFactoryMethod.getInstance();
            Assert.NotNull(instance);
        }

        [Test]
        public void testGetTableItem()
        {
            TableFactoryMethod instance = TableFactoryMethod.getInstance();

            Assert.True(instance.getTableItem(TableEle.TABLE_DEF) is TableDefinition);
            Assert.True(instance.getTableItem(TableEle.TH) is TableHeader);
            Assert.True(instance.getTableItem(TableEle.TD) is TableCol);
            Assert.True(instance.getTableItem(TableEle.TF) is TableFooter);
            Assert.True(instance.getTableItem(TableEle.TD) is TableCol);

            Assert.Null(instance.getTableItem(null));
        }
    }
}