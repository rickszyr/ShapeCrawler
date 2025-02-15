using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.Linq;
using FluentAssertions;
using NUnit.Framework;
using ShapeCrawler.Tests.Shared;
using ShapeCrawler.Tests.Unit.Helpers;
using ShapeCrawler.Tests.Unit.Helpers.Attributes;
using Xunit;

namespace ShapeCrawler.Tests.Unit;

[SuppressMessage("Usage", "xUnit1013:Public method should be marked as test")]
public class ParagraphTests : SCTest
{
    [Test]
    public void IndentLevel_Setter_sets_indent_level()
    {
        // Act
        var pres = SCPresentation.Create();
        var addedShape = pres.Slides[0].Shapes.AddRectangle(100,100, 500, 100);
        var paragraph = addedShape.TextFrame!.Paragraphs.Add();
        paragraph.Text = "Test";
        
        // Act
        paragraph.IndentLevel = 2;

        // Assert
        paragraph.IndentLevel.Should().Be(2);
    }
    
    [Test]
    public void Bullet_FontName_Getter_returns_font_name()
    {
        // Arrange
        var pptx = GetInputStream("002.pptx");
        var pres = SCPresentation.Open(pptx);
        var shapes = pres.Slides[1].Shapes;
        var shape3Pr1Bullet = ((IAutoShape)shapes.First(x => x.Id == 3)).TextFrame.Paragraphs[0].Bullet;
        var shape4Pr2Bullet = ((IAutoShape)shapes.First(x => x.Id == 4)).TextFrame.Paragraphs[1].Bullet;

        // Act
        var shape3BulletFontName = shape3Pr1Bullet.FontName;
        var shape4BulletFontName = shape4Pr2Bullet.FontName;

        // Assert
        shape3BulletFontName.Should().BeNull();
        shape4BulletFontName.Should().Be("Calibri");
    }

    [Test]
    public void Bullet_Type_Getter_returns_bullet_type()
    {
        // Arrange
        var pptx = GetInputStream("002.pptx");
        var pres = SCPresentation.Open(pptx);
        var shapeList = pres.Slides[1].Shapes;
        var shape4 = shapeList.First(x => x.Id == 4);
        var shape5 = shapeList.First(x => x.Id == 5);
        var shape4Pr2Bullet = ((IAutoShape)shape4).TextFrame.Paragraphs[1].Bullet;
        var shape5Pr1Bullet = ((IAutoShape)shape5).TextFrame.Paragraphs[0].Bullet;
        var shape5Pr2Bullet = ((IAutoShape)shape5).TextFrame.Paragraphs[1].Bullet;

        // Act
        var shape5Pr1BulletType = shape5Pr1Bullet.Type;
        var shape5Pr2BulletType = shape5Pr2Bullet.Type;
        var shape4Pr2BulletType = shape4Pr2Bullet.Type;

        // Assert
        shape5Pr1BulletType.Should().Be(SCBulletType.Numbered);
        shape5Pr2BulletType.Should().Be(SCBulletType.Picture);
        shape4Pr2BulletType.Should().Be(SCBulletType.Character);
    }

    [Test]
    public void Alignment_Setter_updates_text_alignment_of_table_Cell()
    {
        // Arrange
        var pres = SCPresentation.Create();
        var table = pres.Slides[0].Shapes.AddTable(10, 10, 2, 2);
        var cellTextFrame = table.Rows[0].Cells[0].TextFrame;
        cellTextFrame.Text = "some-text";
        var paragraph = cellTextFrame.Paragraphs[0];
        
        // Act
        paragraph.Alignment = SCTextAlignment.Center;
        
        // Assert
        paragraph.Alignment.Should().Be(SCTextAlignment.Center);
        var errors = PptxValidator.Validate(pres);
        errors.Should().BeEmpty();
    }

    [Test]
    public void Paragraph_Bullet_Type_Getter_returns_None_value_When_paragraph_doesnt_have_bullet()
    {
        // Arrange
        var pptx = GetInputStream("001.pptx");
        var pres = SCPresentation.Open(pptx);
        var autoShape = pres.Slides[0].Shapes.GetById<IAutoShape>(2);
        var bullet = autoShape.TextFrame.Paragraphs[0].Bullet;

        // Act
        var bulletType = bullet.Type;

        // Assert
        bulletType.Should().Be(SCBulletType.None);
    }

    [Test]
    public void Paragraph_BulletColorHexAndCharAndSizeProperties_ReturnCorrectValues()
    {
        // Arrange
        var pres2 = SCPresentation.Open(GetInputStream("002.pptx"));
        var shapeList = pres2.Slides[1].Shapes;
        var shape4 = shapeList.First(x => x.Id == 4);
        var shape4Pr2Bullet = ((IAutoShape)shape4).TextFrame.Paragraphs[1].Bullet;

        // Act
        var bulletColorHex = shape4Pr2Bullet.ColorHex;
        var bulletChar = shape4Pr2Bullet.Character;
        var bulletSize = shape4Pr2Bullet.Size;

        // Assert
        bulletColorHex.Should().Be("C00000");
        bulletChar.Should().Be("'");
        bulletSize.Should().Be(120);
    }
        
    [Test]
    public void Paragraph_Text_Setter_updates_paragraph_text_and_resize_shape()
    {
        // Arrange
        var pptxStream = GetInputStream("autoshape-case003.pptx");
        var pres = SCPresentation.Open(pptxStream);
        var shape = pres.Slides[0].Shapes.GetByName<IAutoShape>("AutoShape 4");
        var paragraph = shape.TextFrame.Paragraphs[0];
            
        // Act
        paragraph.Text = "AutoShape 4 some text";

        // Assert
        shape.Height.Should().Be(46);
        shape.Y.Should().Be(148);
    }

    [Test]
    public void Text_Setter_sets_paragraph_text_in_New_Presentation()
    {
        // Arrange
        var pres = SCPresentation.Create();
        var slide = pres.Slides[0];
        var shape = slide.Shapes.AddRectangle(10, 10, 10, 10);
        var paragraph = shape.TextFrame!.Paragraphs[0];

        // Act
        paragraph.Text = "test";

        // Assert
        paragraph.Text.Should().Be("test");
        var errors = PptxValidator.Validate(slide.Presentation);
        errors.Should().BeEmpty();
    }
    
    [Test]
    public void Text_Setter_sets_paragraph_text_for_grouped_shape()
    {
        // Arrange
        var pptx = TestHelper.GetStream("autoshape-case003.pptx");
        var pres = SCPresentation.Open(pptx);
        var shape = pres.Slides[0].Shapes.GetByName<IGroupShape>("Group 1").Shapes.GetByName<IAutoShape>("Shape 1");
        var paragraph = shape.TextFrame!.Paragraphs[0];
        
        // Act
        paragraph.Text = $"Safety{Environment.NewLine}{Environment.NewLine}{Environment.NewLine}";
        // SaveResult(pres);
        
        // Assert
        paragraph.Text.Should().BeEquivalentTo($"Safety{Environment.NewLine}{Environment.NewLine}{Environment.NewLine}");
        TestHelper.ThrowIfPresentationInvalid(pres);
    }
    
    [Test]
    public void Paragraph_Text_Getter_returns_paragraph_text()
    {
        // Arrange
        var textBox1 = ((IAutoShape)SCPresentation.Open(GetInputStream("008.pptx")).Slides[0].Shapes.First(sp => sp.Id == 37)).TextFrame;
        var textBox2 = ((ITable)SCPresentation.Open(GetInputStream("009_table.pptx")).Slides[2].Shapes.First(sp => sp.Id == 3)).Rows[0].Cells[0]
            .TextFrame;

        // Act
        string paragraphTextCase1 = textBox1.Paragraphs[0].Text;
        string paragraphTextCase2 = textBox1.Paragraphs[1].Text;
        string paragraphTextCase3 = textBox2.Paragraphs[0].Text;

        // Assert
        paragraphTextCase1.Should().BeEquivalentTo("P1t1 P1t2");
        paragraphTextCase2.Should().BeEquivalentTo("p2");
        paragraphTextCase3.Should().BeEquivalentTo("0:0_p1_lvl1");
    }

    [Test]
    public void ReplaceText_finds_and_replaces_text()
    {
        // Arrange
        var pptxStream = GetInputStream("autoshape-case003.pptx");
        var pres = SCPresentation.Open(pptxStream);
        var paragraph = pres.Slides[0].Shapes.GetByName<IAutoShape>("AutoShape 3").TextFrame!.Paragraphs[0];
            
        // Act
        paragraph.ReplaceText("Some text", "Some text2");

        // Assert
        paragraph.Text.Should().BeEquivalentTo("Some text2");
        var errors = PptxValidator.Validate(pres);
        errors.Should().BeEmpty();
    }

    [Test]
    public void Paragraph_Portions_counter_returns_number_of_text_portions_in_the_paragraph()
    {
        // Arrange
        var textFrame = SCPresentation.Open(GetInputStream("009_table.pptx")).Slides[2].Shapes.GetById<IAutoShape>(2).TextFrame;

        // Act
        var portions = textFrame.Paragraphs[0].Portions;

        // Assert
        portions.Should().HaveCount(2);
    }
    
    [Test]
    public void Portions_Add()
    {
        // Arrange
        var pptx = GetInputStream("autoshape-case001.pptx");
        var pres = SCPresentation.Open(pptx);
        var shape = pres.SlideMasters[0].Shapes.GetByName<IAutoShape>("AutoShape 1");
        var paragraph = shape.TextFrame!.Paragraphs.Add();
        var expectedPortionCount = paragraph.Portions.Count + 1;
        
        // Act
        paragraph.Portions.AddText(" ");
        
        // Assert
        paragraph.Portions.Count.Should().Be(expectedPortionCount);
    }
}