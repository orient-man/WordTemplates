namespace WordTemplates
{
    public interface ITemplateNavigator
    {
        /// <summary>
        /// Sets value of field (and its synonyms)
        /// </summary>
        void SetField(string field, object value);

        /// <summary>
        /// Sets values in table assuming 1 existing row
        /// </summary>
        void SetFields(string[] fields, object[][] values);
    }
}