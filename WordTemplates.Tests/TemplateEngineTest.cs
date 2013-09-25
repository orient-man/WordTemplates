using System.IO;
using NUnit.Framework;

namespace WordTemplates.Tests
{
    [TestFixture]
    [Category("Integration")]
    public class TemplateEngineTest
    {
        private const string TemplateFile = @"template.dotx";

        [Test]
        public void SmokeTest()
        {
            const string outputFile = @"filled.doc";
            WordHelper.FillTemplate(
                TemplateFile,
                new[]
                    {
                        new TemplateField { Name = "CompanyName", Value = "Milena-Tech" },
                        new TemplateField { Name = "FirstName", Value = "Marcin" },
                        new TemplateField { Name = "Surname", Value = "Malinowski" },
                        new TemplateField { Name = "Checkbox", Value = true },
                        new TemplateField
                            {
                                Name = "ArticleName",
                                Value = "iPhone 5S",
                                Iteration = 0,
                                Group = "Articles"
                            },
                        new TemplateField
                            {
                                Name = "Units",
                                Value = "10",
                                Iteration = 0,
                                Group = "Articles"
                            },
                        new TemplateField
                            {
                                Name = "ArticleName",
                                Value = "iPhone 5C",
                                Iteration = 1,
                                Group = "Articles"
                            },
                        new TemplateField
                            {
                                Name = "Units",
                                Value = "2",
                                Iteration = 1,
                                Group = "Articles"
                            },
                        new TemplateField
                            {
                                Name = "ArticleName",
                                Value = "iPad Mini",
                                Iteration = 2,
                                Group = "Articles"
                            },
                        new TemplateField
                            {
                                Name = "Units",
                                Value = "999",
                                Iteration = 2,
                                Group = "Articles"
                            }
                    },
                outputFile);
        }
    }
}