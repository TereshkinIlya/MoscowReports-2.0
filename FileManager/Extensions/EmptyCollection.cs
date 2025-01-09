namespace FileManager.Extensions
{
    public static class EmptyCollection
    {
        public static bool IsNullOrEmptyCollection(this IEnumerable<string> collection)
        {
            return collection == null || !collection.Any();
        }
    }
}
