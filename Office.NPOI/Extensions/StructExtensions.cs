namespace Ideal.Office.Excel
{
    using System;
    using System.Text;

    public static class StructExtensions
    {
        /// <summary>
        /// 可转换A~Z为相关的数据
        /// </summary>
        public static int ToNumber(this string value) {
            byte[] array = Encoding.ASCII.GetBytes(value[0].ToString().ToUpper());
            int asciicode = (int)(array[0]);
            int tempValue = 0;
            if (int.TryParse(value, out tempValue))
            {
                //return tempValue;
            }
            else if (asciicode > 64 && asciicode < 91)
            {
                tempValue = asciicode - 64;
            }
            return tempValue;
        }

        /// <summary>
        /// 将Object转换为Double
        /// </summary>
        public static double ToDouble(this object value)
        {
            return Convert.ToDouble(value);
        }

        /// <summary>
        /// 将Double转换为Decimal
        /// </summary>
        public static Decimal ToDecimal(this double value)
        {
            Decimal tempValue = 0L;
            if (Decimal.TryParse(value.ToString(), out tempValue))
            {
                //return tempValue;
            }
            return tempValue;
        }

        /// <summary>
        /// 截取最长的字符串
        /// </summary>
        public static string SubString(this string value, int maxLength)
        {
            if (string.IsNullOrEmpty(value))
                return "";

            if (value.Length < maxLength)
                return value;

            return value.Substring(0, maxLength);
        }

    }
}
