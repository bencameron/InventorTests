using Xunit;

namespace InventorTests
{
    public class Tests
    {
        [Fact]
        public void CanGetInventorInstance()
        {
            var inventorApp = InventorAppUtils.GetInventorInstance();
            Assert.NotNull(inventorApp);
        }
    }
}
