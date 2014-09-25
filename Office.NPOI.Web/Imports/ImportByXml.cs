namespace Ideal.Office.Web
{
    using System;
    using System.Collections.Generic;
    using System.Web;
    using System.Linq;

    using Configuration;

    using Ideal.Office.Excel;

    public partial class Import
    {
        private static string MapPath(string path)
        {
            return HttpContext.Current.Server.MapPath(path);
        }

        public static IDictionary<string, ExcelEntity> GetXmlConfig()
        {
            var templates = new Dictionary<string, ExcelEntity>();
            var dic = new Dictionary<string, ExcelType> { 
                    {"Excel2003", ExcelType.Excel2003 },
                    {"Excel2007", ExcelType.Excel2007 },
                    {"Excel2010", ExcelType.Excel2010 },
                    {"Excel2013", ExcelType.Excel2013 }
                };

            var config = CatalogsSection.GetInstance();
            if (config != null && config.Catalogs != null)
            {

                config.Catalogs
                      .Cast<CatalogElement>()
                      .ForEach(item =>
                      {
                          var c = item.Template;
                          var contsflag = new ConstFlag { 
                            Consts = c.Consts.Cast<System.Configuration.NameValueConfigurationElement>().ToList()
                          };
                          

                          if (!string.IsNullOrEmpty(c.Path))
                          {
                              string path = c.Path;
                              if (path.StartsWith("~"))
                                  path = MapPath(path);

                              var param = new List<Tuple<SheetFlag, ConstFlag, SqlFlag>>();
                              c.Cast<SqlItemElementCollection>()
                               .ForEach(sql =>
                               {
                                   var s = sql.Start.Split(',');
                                   var p = string.IsNullOrEmpty(sql.Param) ? new string[0] : sql.Param.Split(',');
                                   var fs = sql.Formula.Split(';');
                                   var f = new List<FormulaEntity>();
                                   fs.ForEach(formula =>
                                   {
                                       if (!string.IsNullOrEmpty(formula))
                                       {
                                           var formulaItem = formula.Split('|');
                                           f.Add(new FormulaEntity
                                           {
                                               Postion = formulaItem[0],
                                               ColNo = int.Parse(formulaItem[1]),
                                               FormulaText = formulaItem[2]
                                           });
                                       }
                                   });

                                   var childs = new Dictionary<string, string>();
                                   sql.Cast<AddItemElement>()
                                      .ForEach(add =>
                                      {
                                          childs.Add(add.Key, add.Value);
                                      });

                                   param.Add(Tuple.Create<SheetFlag, ConstFlag, SqlFlag>(
                                             new SheetFlag
                                             {
                                                 SheetIndex = sql.SheetIndex
                                              ,
                                                 Row = s[0].ToNumber() - 1
                                              ,
                                                 Col = s[1].ToNumber() - 1
                                              ,
                                                 MergedColPrimaryKey = sql.MergedColPrimaryKey
                                              ,
                                                 MergedCols = string.IsNullOrEmpty(sql.MergedCols) ? new int[0] : Array.ConvertAll<string, int>(sql.MergedCols.Split(','), m1 => { return int.Parse(m1); })
                                              ,
                                                 MergedRowPrimaryKey = sql.MergedRowPrimaryKey
                                              ,
                                                 MergedRows = string.IsNullOrEmpty(sql.MergedRows) ? new int[0] : Array.ConvertAll<string, int>(sql.MergedRows.Split(','), m2 => { return int.Parse(m2); })
                                              ,
                                                 Formulas = f
                                              ,
                                                 Childs = childs
                                             }
                                           , contsflag
                                           , new SqlFlag { Query = sql.Value, Param = p }));

                               });

                              var entity = new ExcelEntity
                              {
                                  Name = c.Name,
                                  Path = path,
                                  Type = dic[c.Type],
                                  Param = param
                              };

                              templates.Add(c.Name, entity);
                          }
                      });
            }
            return templates;
        }

        public static IDictionary<string, ExcelEntity> GetImportConfig()
        {
            var config = EntitySection.GetInstance();
            if (config != null && config.Entity != null)
            {
                config.Entity
                      .Cast<EntityElement>()
                      .ForEach(item => {

                          var id = item.Id;

                          var cls = item.Class.Cast<ClassElement>().ToList();


                          //var cols = item.Columns.Cast<ColumnElementCollection>().Cast<ColumnElement>().ToList();
                          //var col = cols.FirstOrDefault();

                          var cols = item.Columns.Cast<ColumnElementCollection>().ToList();
                          var cols1 = cols.FirstOrDefault().Cast<ColumnElement>().ToList();

                          //Export.Deonw(cols);

                      });
            }

            return null;
        }
    }
}
