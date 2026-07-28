using System.Text.Json;

namespace ShivaNegar.Models
{
    internal class UpdateResultModel
    {
        private Results result;
        private JsonElement content;

        internal UpdateResultModel(Results result, JsonElement content)
        {
            Result = result;
            Content = content;
        }

        internal JsonElement Content { get => content; set => content = value; }
        internal Results Result { get => result; set => result = value; }

        internal enum Results
        {
            ServerProblem,
            ServerNotAvailable,
            ServerResult_NoUpdateAvailable,
            ServerResult_UpdateOptional,
            ServerResult_UpdateMandatory
        }
    }
}
