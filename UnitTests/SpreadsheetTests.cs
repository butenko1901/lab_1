using System.Numerics;
using MyExcelMAUIApp.Parsing;
using Xunit;

namespace UnitTests;

public class SpreadsheetTests
{
    private static CellResult Calculate(string expression)
    {
        var sheet = new Spreadsheet(5, 5);
        sheet.SetExpression(1, 1, expression);
        return sheet.Evaluate(1, 1);
    }

    [Fact]
    public void ModAndDiv_GiveRemainderAndQuotient()
    {
        var mod = Calculate("17 mod 5");
        var div = Calculate("17 div 5");

        Assert.Equal(new BigInteger(2), (BigInteger)mod.Value!);
        Assert.Equal(new BigInteger(3), (BigInteger)div.Value!);
    }

    [Fact]
    public void LogicalExpression_NotAndComparison_IsTrue()
    {
        var result = Calculate("not (3 > 5) and 2 = 2");

        Assert.False(result.IsError);
        Assert.True((bool)result.Value!);
    }

    [Fact]
    public void CellReference_UsesValueOfOtherCell()
    {
        var sheet = new Spreadsheet(5, 5);
        sheet.SetExpression(1, 1, "mmax(4, 9, 2)");
        sheet.SetExpression(2, 1, "A1 mod 4");

        var result = sheet.Evaluate("A2");

        Assert.Equal(new BigInteger(1), (BigInteger)result.Value!);
    }

    [Fact]
    public void DivisionByZero_ReturnsErrorWithPosition()
    {
        var result = Calculate("5 div 0");

        Assert.True(result.IsError);
        Assert.Equal("Ділення на нуль", result.Error);
        Assert.Equal(2, result.ErrorPosition);
    }
}
