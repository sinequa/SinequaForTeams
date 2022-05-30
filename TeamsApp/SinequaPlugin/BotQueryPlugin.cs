using Sinequa.Common;
using Sinequa.Engine.Client;
using Sinequa.Plugins;
using Sinequa.Search;
using System;
using System.Text;

namespace TeamsMessagingExtensionsSearch.SinequaPlugin
{
    /*
     * We override the standard queryPlugin to bypass the use of SearchSession
     */
    public class BotQueryPlugin : QueryPlugin
    {

        //internal IEngineClient InternalClient { get; set; }
        public string userid { get; set; }
        public string domain { get; set; }
        public override void AddSelect(StringBuilder sb)
        {
            Str.Add(sb, "select ");
            var startIndex = sb.Length;

            AddStandardColumns(sb);
            AddRelevanceColumns(sb);
            AddQueryColumns(sb);
            AddFormatTextColumn(sb);

            SqlParts.Select = sb.ToString(startIndex, sb.Length - startIndex);
        }


        public override bool DoQuery()
        {
            //if (AggregationsOnly && !HaveActiveAggregations())
            //{
            //    WritePredefinedAggregations();
            //    return true;
            //}
            var statementsOffset = Statements.Count;
            string sql = ToSql();
            Statements.Add(sql);
            if (Str.IsEmpty(sql)) return false;
            try
            {
                using (Cursor cursor = Client.ExecCursor(sql, SqlTimeout))
                {
                    WriteResults(null, cursor);
                }
                //if (QueryAnalysis != null) //We will only have the query analysis when the queryintents have been evaluated.
                //{
                //    QueryIntent.MatchActionAfterSearch(this, QueryAnalysis, jsonResult: Response);
                //}
                return true;
            }
            catch (Exception e)
            {
                //LogQueryExecutionError($"{nameof(QueryPlugin)}.DoQuery", e);
                return false;
            }
        }



        // Details is an array of detail objects. Each detail object has an array of statements (because of fielded search)
        // and the attributes associated with the principal request. There is a detail object for each query (tab or main)
        // providing per-query processing time information.
        public override bool DoTabQuery()
        {
            return true;
        }


        public override void AddQueryColumns(StringBuilder sb)
        {
            string columnsToAddStr = !string.IsNullOrEmpty(QueryColumns) ? QueryColumns : Query.Columns;
            //columnsToAddStr = FilterColumnsExcludedFromSearch(columnsToAddStr);
            if (!Str.IsEmpty(columnsToAddStr) && !Str.BeginWith(columnsToAddStr, ","))
            {
                Str.Add(sb, ',');
            }
            Str.Add(sb, columnsToAddStr);
            var columnsToAdd = ListStr.ListFromStr(columnsToAddStr, ',', ListStr.FromStrFlags.Trim | ListStr.FromStrFlags.NoEmpty);
            SqlParts?.Columns.Add(columnsToAdd);
        }

        public override void AddRights(StringBuilder sb)
        {

            //TODO Maybe can do better for rights ?
            sb.Append(And()).Append("(CACHE (CHECKACLS('accesslists=\"accesslist1,accesslist2\",deniedlists=\"deniedlist1\",security_column=\"false\"') FOR ('"+domain+"|"+userid+"')))");
        }

        protected override bool IsColumnExcludedFromSearch(string name)
        {
            return false;
        }

        public override void WriteResults(Json detail, Cursor cursor)
        {
            Response.Set("id", Sys.NewId()); // results id (for audit)
            Response.Set("queryName", Query.Name);

            WriteRecords(cursor);

        }
    }
}
