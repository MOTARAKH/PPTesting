
// By Mohamad Tarekh :: tarakhmohamad2002@gmail.com
// !!!!!
// download NuGet Package=> 'Syncfusion.Presentation.Net.Core'
// !!!!!
using Syncfusion.Drawing;
using Syncfusion.Presentation;

//Open an existing PowerPoint presentation

IPresentation pptxDoc = Presentation.Open(new FileStream(path: "C:\\Users\\MHTH\\source\\repos\\Testing\\Testing\\firstpp.pptx", FileMode.Open));

//Gets the first slide from the PowerPoint presentation

ISlide slide = pptxDoc.Slides[0];

var count = slide.Shapes.Count;
//Gets the first shape of the slide

IFont font;

// Assuming you want to change the position of the first shape on the slide

// Set the new left and top positions (in points)
float newLeft = 210; // Adjust this value as needed
float newTop = 136; // Adjust this value as needed
float newLeft3 = 385;
float newLeft4 = 560;
// Assuming you want to change the height and width of the first shape on the slide
IShape shape1= (IShape)slide.Shapes[1];
shape1.TextBody.Text = "Begin";
shape1.TextBody.Paragraphs[0].TextParts[0].Font.Color = ColorObject.White;
IShape shape2 = (IShape)slide.Shapes[2];
shape2.TextBody.Text = "Step1";
shape2.TextBody.Paragraphs[0].TextParts[0].Font.Color = ColorObject.White;
IShape shape3= (IShape)slide.Shapes[3];
shape3.TextBody.Text = "Step2";
shape3.TextBody.Paragraphs[0].TextParts[0].Font.Color = ColorObject.White;
IShape shape4= (IShape)slide.Shapes[4];
shape4.TextBody.Text = "Done";
shape4.TextBody.Paragraphs[0].TextParts[0].Font.Color = ColorObject.White;

// Set the new height and width (in points)
float newHeight = 112; // Adjust this value as needed
float newWidth = 222; // Adjust this value as needed

shape1.Height = newHeight;
shape1.Width = newWidth;


shape2.Height = newHeight;
shape2.Width = newWidth;
shape2.Left = newLeft;
shape2.Top = newTop;


shape3.Height = newHeight;
shape3.Width = newWidth;
shape3.Top=newTop;
shape3.Left = newLeft3;

shape4.Height = newHeight;
shape4.Width = newWidth;
shape4.Top=newTop;
shape4.Left = newLeft4;


slide.Shapes[0].Left = 300;
IShape ss = slide.Shapes[0] as IShape;
ss.TextBody.Text = "OutPut Slide";

for (int i = 0; i < 5; i++)
{
	IShape ishape = (IShape)slide.Shapes[5];
	var iSlideItem = slide.Shapes[i].GetType();

	if (ishape.TextBody.Text == "Step 1" || ishape.TextBody.Text == "Step 2" || ishape.TextBody.Text == "Begin" || ishape.TextBody.Text == "Done")
	{
		slide.Shapes.Remove(ishape);
		count--;
	}
	
	Console.WriteLine(iSlideItem.Name);

}



for (int i = 0; i < 4; i++)
{
	
	IShape shape = slide.Shapes[i+5] as IShape;
	var iSlideItem = slide.Shapes[i].GetType();
	for(int j=0; j < 4; j++)
	{
		IParagraph paragraph = shape.TextBody.Paragraphs[j];
		//Retrieves the first TextPart of the shape.

		ITextPart textPart = paragraph.TextParts[0];
		textPart.Font.Underline = TextUnderlineType.None;
		textPart.Font.Italic = false;
		textPart.Font.Bold = false;
		paragraph.ListFormat.Type=ListType.Bulleted;
		paragraph.ListFormat.BulletCharacter = Convert.ToChar(183);
		paragraph.ListFormat.Size = 70;
		paragraph.ListFormat.FontName = "Symbol";
		

	}




	Console.WriteLine(iSlideItem.Name);
}


//Save the PowerPoint presentation as stream

using (var stream = new FileStream("C:\\Users\\MHTH\\source\\repos\\Testing\\Testing\\last_out_343.pptx", FileMode.Create))
{
	pptxDoc.Save(stream);
	stream.Position = 0;
}

pptxDoc.Close();