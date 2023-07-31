/* 
pareto creator c# - considers RANK() window function

Compatabile with Tabular Editor 2 and 3

How to use:
1. Set your user defined strings and save macro
2. Select measures for your pareto analysis execute the macro

Pareto in a statistical analysis tool primarily used in quality control scenarios (see: 80-20 rule)

Output: 3 measures per selected measure AND column in dimension table
    1. Ranking measure (ranks categories by selected measure, descending)
    2. Percentage Line measure (caclulates running total % across categories)
    3. Bin color measure (assigns color to columns based on being in the 80 or 20 % portion of the sample
*/

// User defined strings/variables
string dimTblNm = "Order Status";
string calcTblNm = "Calculations";
string EightyBinColor = "#123456";  // hex color string for the 80% grouping (majority of data points)
string TwentyBinColor = "#ececec";  // hex color string for the 20% grouping (minority of data points)

// set variables
var dimTbl = Model.Tables[dimTblNm];
var calcTbl = Model.Tables[calcTblNm];

foreach ( var m in Selected.Measures )
{
    string m_dx = m.DaxObjectName;
    foreach ( var c in Model.AllColumns.Where(c => 
        c.Table.Name == dimTbl.Name &&
        c.DataType == DataType.String &&
        c.IsHidden == false
    ))
    {
        string vt = "";
        if ( c.SortByColumn == null )
        {   
            vt = $"ALLSELECTED ( {c.DaxObjectFullName} )";
        }
        else
        {
            vt = $"ALLSELECTED ( {c.DaxObjectFullName}, {c.SortByColumn.DaxObjectFullName} )";
        }
        string vt_ac = $"ADDCOLUMNS( {vt}, \"@RankingAmount\", {m_dx} )";
        
        string displayFolder = $"Pareto\\{m.Name}\\{c.Name}";
        Measure RankM = c.Table.AddMeasure(
            $"Rank {c.Name} by {m.Name}", 
            $"\nvar _baseTbl = {vt_ac}\nvar _rnk = \nRANK ( \n\tDENSE, \n\t_baseTbl, \n\tORDERBY( [@RankingAmount], DESC, {c.DaxObjectFullName}, ASC BLANKS LAST )\n) \nRETURN \n\t_rnk" ,
            displayFolder
        );

        Measure PctLine = c.Table.AddMeasure(
            $"{m.Name} + Pareto Pct Line ({c.Name})",
            $"VAR __numerator = VAR __topNTbl = TOPN ( {RankM.DaxObjectName}, {vt}, {RankM.DaxObjectName}, ASC ) RETURN SUMX ( __topNTbl, {m_dx} ) VAR __denominator = CALCULATE ( SUMX ( VALUES ( {c.DaxObjectFullName} ), {m_dx} ), ALLSELECTED ( {c.Table.DaxObjectName} ) ) VAR _logicalDisplay = DIVIDE( 1, NOT ISBLANK( {m_dx} ) ) VAR __result = DIVIDE ( __numerator, __denominator ) RETURN DIVIDE( __result, _logicalDisplay )",
            displayFolder
        );
        PctLine.FormatDax();
        PctLine.FormatString = "0%";
        
        Measure BinColors = c.Table.AddMeasure(
            $"{m.Name} Bin Colors ({c.Name})",
            $"\nvar _eighty = \"{EightyBinColor}\"\nvar _twenty = \"{TwentyBinColor}\"\nvar _meas = {PctLine.DaxObjectName}\nvar _rnk = {RankM.DaxObjectName}\nRETURN\n\tSWITCH( TRUE(), ISBLANK(_meas), BLANK(), _rnk = 1, _eighty, _meas > 0.8 , _twenty, _eighty )   ",
            displayFolder
        );
      
    }
}