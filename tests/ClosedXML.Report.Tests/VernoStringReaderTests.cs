using System;
using System.Globalization;
using System.Threading;
using ClosedXML.Report.Utils;
using FluentAssertions;
using Xunit;

namespace ClosedXML.Report.Tests
{
    public class VernoStringReaderTests
    {
        [Fact]
        public void ReadTest1()
        {
            var reader = new VernoStringReader("Lorem  ipsum  325  dolor sit amet, 'Phasellus in nunc ac sem aliquam tempus.' " +
                                                    " \"consectetur=123 \" val=<test>  adipiscing elit.");
            reader.ReadWord().Should().Be("Lorem");
            reader.ReadWord().Should().Be("ipsum");
            reader.ReadInt().Should().Be(325);
            reader.ReadWhile(ch => ch != ',').ToString().Should().Be("dolor sit amet");
            reader.ReadInBrackets('\'').ToString().Should().Be("Phasellus in nunc ac sem aliquam tempus.");
            reader.ReadTo(new[] {"="}, true);
            reader.ReadRegex(@"\<(.+?)\>").ToString().Should().Be("test");
            reader.SkipSpace();
            reader.ReadToEnd().ToString().Should().Be("adipiscing elit.");
            reader.LastChar.Should().Be('.');
        }

        [Fact]
        public void ReadTest2()
        {
            var reader = new VernoStringReader("[val=65];[tag\\val2=option;val3=25.65]");
            var r1 = reader.ReadInBrackets('[', ']');
            r1.ToString().Should().Be("val=65");
            r1.ReadWord().Should().Be("val");
            r1.ReadChar().Should().Be('=');
            r1.ReadInt().Should().Be(65);
            reader.ReadChar().Should().Be(';');
            var r2 = reader.ReadInBrackets('[', ']');
            r2.ToString().Should().Be("tag\\val2=option;val3=25.65");
            r2.ReadTo('\\').ToString().Should().Be("tag");
            r2.LastChar.Should().Be('\\');
        }

        [Fact]
        void ReadInBracketsTests()
        {
            var reader = new VernoStringReader("Sed id urna ac quam congue (venenatis.");
            reader.ReadInBrackets('(', ')').ToString().Should().Be("venenatis.");

            reader = new VernoStringReader("35*(2+2*(43-x))+23*y");
            reader.ReadInt().Should().Be(35);
            reader.Peek().Should().Be('*');
            var inreader = reader.ReadInBrackets('(', ')');
            inreader.ToString().Should().Be("2+2*(43-x)");
            inreader.ReadTo('*');
            inreader.ReadInBrackets('(', ')').ToString().Should().Be("43-x");
        }

        [Fact]
        public void ReadArrayTest()
        {
            var reader = new VernoStringReader("val1; val2; val3; val4");
            var array = reader.ReadArray(";");
            array.Length.Should().Be(4);
            array.Should().ContainInOrder("val1", "val2", "val3", "val4");

            reader = new VernoStringReader("'Sed pharetra feugiat ante.';'Suspendisse;eget';'nulla vitae arcu';'interdum ; scelerisque.");
            array = reader.ReadArray(";");
            array.Length.Should().Be(4);
            array.Should().ContainInOrder("Sed pharetra feugiat ante.", "Suspendisse;eget", "nulla vitae arcu", "interdum ; scelerisque.");

            reader = new VernoStringReader("10, 20, 30, 40");
            reader.ReadArray<int>(",").Should().BeEquivalentTo(10, 20, 30, 40);
        }

        [Fact]
        public void ReadNumbersTest()
        {
            var reader = new VernoStringReader("10 15.4 23.4 54,2 sd 0-3 -4");
            reader.ReadInt().Should().Be(10);
            reader.ReadInt().Should().Be(15);
            reader.ReadDouble().Should().Be(4d);
            reader.ReadDouble().Should().Be(23.4);
            reader.ReadDouble().Should().Be(54.2);
            reader.ReadInt(s =>
            {
                s.Should().Be("sd");
                return 9;
            }).Should().Be(9);
            reader.ReadInt().Should().Be(0);
            reader.ReadInt().Should().Be(-3);
            reader.ReadInt().Should().Be(-4);
        }

        [Fact]
        public void ReadDateTimeTest()
        {
            //TODO make tests culture-independent
            Thread.CurrentThread.CurrentCulture = CultureInfo.CreateSpecificCulture("ru-RU");
            var reader = new VernoStringReader("My birthday is 28.12.1979 (Дек 15, old style)", CultureInfo.GetCultureInfo("ru-RU"));
            reader.ReadTo("s");
            reader.ReadDateTime().Should().Be(DateTime.Parse("28.12.1979"));
            reader.ReadInBrackets('(', ')').ReadDateTime("MMM dd").Should().Be(new DateTime(DateTime.Today.Year, 12, 15));
        }

        [Fact]
        public void ReadNamedValuesTest()
        {
            var reader = new VernoStringReader("name1='some string;value=23';'name 2'=35;name_3=68.4");
            var values = reader.ReadNamedValues(";", "=");
            values.Length.Should().Be(3);
            values[0].Key.Should().Be("name1");
            values[0].Value.Should().Be("some string;value=23");
            values[1].Key.Should().Be("name 2");
            values[1].Value.Should().Be("35");
            values[2].Key.Should().Be("name_3");
            values[2].Value.Should().Be("68.4");
        }
    }
}