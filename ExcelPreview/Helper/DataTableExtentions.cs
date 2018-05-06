using System;
using System.Collections.Generic;
using System.Data;
using System.Reflection;

namespace ExcelPreview.Helper
{
    public static class DataTableMapping
    {
        // function that set the given object from the given data row
        public static void SetItemFromRow<T>(T item, DataRow row)
            where T : new()
        {
            // go through each column
            foreach (DataColumn c in row.Table.Columns)
            {
                // find the property for the column
                PropertyInfo p = item.GetType().GetProperty(c.ColumnName);

                // if exists, set the value
                if (p != null && row[c] != DBNull.Value)
                {
                    p.SetValue(item, row[c], null);
                }

            }
        }

        // function that creates an object from the given data row
        public static T CreateItemFromRow<T>(DataRow row)
            where T : new()
        {
            // create a new object
            T item = new T();

            // set the item
            SetItemFromRow(item, row);

            // return 
            return item;
        }

        // function that creates a list of an object from the given data table
        public static List<T> CreateListFromTableForAll<T>(DataTable tbl)
            where T : new()
        {
            // define return list
            List<T> lst = new List<T>();

            // alocal variable
            int count = 0;

            // go through each row
            foreach (DataRow r in tbl.Rows)
            {
                // add to the list
                lst.Add(CreateItemFromRow<T>(r));
                count++;
            }

            // return the list
            return lst;
        }


        // function that creates a list of an object from the given data table
        public static T CreateObjectFromTableForAll<T>(DataTable tbl)
            where T : new()
        {
            // define return obj
            T obj = new T();
            // go through each row
            foreach (DataRow r in tbl.Rows)
            {
                // add to the obj
                obj = (CreateItemFromRow<T>(r));
            }

            // return the list
            return obj;
        }

        public static DataTable CreateDataTableFromList<T>(List<T> items)
        {
            DataTable dataTable = new DataTable(typeof(T).Name);

            //Get all the properties
            PropertyInfo[] Props = typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance);
            foreach (PropertyInfo prop in Props)
            {
                //Defining type of data column gives proper data table 
                var type = (prop.PropertyType.IsGenericType && prop.PropertyType.GetGenericTypeDefinition() == typeof(Nullable<>) ? Nullable.GetUnderlyingType(prop.PropertyType) : prop.PropertyType);
                //Setting column names as Property names
                dataTable.Columns.Add(prop.Name, type);
            }
            foreach (T item in items)
            {
                var values = new object[Props.Length];
                for (int i = 0; i < Props.Length; i++)
                {
                    //inserting property values to datatable rows
                    values[i] = Props[i].GetValue(item, null);
                }
                dataTable.Rows.Add(values);
            }
            //put a breakpoint here and check datatable
            return dataTable;
        }

    }
}