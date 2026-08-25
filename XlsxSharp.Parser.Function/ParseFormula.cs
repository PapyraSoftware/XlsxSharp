using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Newtonsoft.Json;
using XlsxSharp.Parser.Ast;
using XlsxSharp.Parser.Pratt;

namespace XlsxSharp.Parser.Function
{
    public static class ParseFormula
    {
        [FunctionName("parse-formula")]
        public static Task<IActionResult> Run(
            [HttpTrigger(AuthorizationLevel.Anonymous, "get", Route = null)] HttpRequest req
        )
        {
            ReferenceStyle refStyle =
                req.Query["style"] == "R1C1" ? ReferenceStyle.R1C1 : ReferenceStyle.A1;
            string formulaText = req.Query["formula"];

            JsonSerializerSettings serializerSetting = new()
            {
                Formatting = Formatting.Indented,
                Converters = { new AstNodeConverter(refStyle) },
            };

            try
            {
                AstNode nodes = ParserFactory
                    .Create(new F())
                    .ParseFormula(formulaText, new Ctx(), isR1C1: refStyle == ReferenceStyle.R1C1);
                return Task.FromResult<IActionResult>(
                    new JsonResult(
                        new
                        {
                            formula = formulaText,
                            style = refStyle.ToString(),
                            ast = nodes,
                        },
                        serializerSetting
                    )
                );
            }
            catch (ParsingException e)
            {
                return Task.FromResult<IActionResult>(
                    new JsonResult(
                        new
                        {
                            formula = formulaText,
                            style = refStyle.ToString(),
                            error = e.Message,
                        },
                        serializerSetting
                    )
                    {
                        StatusCode = StatusCodes.Status422UnprocessableEntity,
                    }
                );
            }
        }
    }
}
