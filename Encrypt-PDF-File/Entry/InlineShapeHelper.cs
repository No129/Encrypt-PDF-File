using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Entry
{
    public class InlineShapeHelper
    {

        public Shape Find(string pi_sTitle, Document pi_objTarget)
        {
            Shape objReturn = null;

            foreach (Shape objEachShape in pi_objTarget.Shapes)
            {
                if (objEachShape.Title == pi_sTitle)
                {
                    objReturn = objEachShape;
                }
            }

            return objReturn;
        }
    }
}
