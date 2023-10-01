﻿// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

namespace CallingBotSample.Bots
{
    // Copyright (c) Microsoft Corporation. All rights reserved.
    // Licensed under the MIT License.

    using Microsoft.AspNetCore.Mvc.Rendering;
    using Microsoft.Bot.Connector.Authentication;
    using Microsoft.Graph;
    using Microsoft.Identity.Client;
    using Microsoft.IdentityModel.Clients.ActiveDirectory;
    using Newtonsoft.Json;
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Linq;
    using System.Net.Http;
    using System.Text;
    using System.Text.Json;
    using System.Text.Json.Serialization;
    using System.Threading.Tasks;
    using static System.Formats.Asn1.AsnWriter;
    using AuthenticationResult = Microsoft.Identity.Client.AuthenticationResult;
    using JsonSerializer = System.Text.Json.JsonSerializer;

    public class LLMClient
    {
        const string ENDPOINT = "https://httpqas26-frontend-qasazap-prod-dsm02p.qas.binginternal.com/completions";

        static IEnumerable<string> SCOPES = new List<string>() {
    "api://68df66a4-cad9-4bfd-872b-c6ddde00d6b2/access"
    };

        static IPublicClientApplication app = PublicClientApplicationBuilder.Create("68df66a4-cad9-4bfd-872b-c6ddde00d6b2")
            .WithAuthority("https://login.microsoftonline.com/72f988bf-86f1-41af-91ab-2d7cd011db47")
            .Build();

        private readonly string _llmToken = "eyJ0eXAiOiJKV1QiLCJhbGciOiJSUzI1NiIsImtpZCI6Ii1LSTNROW5OUjdiUm9meG1lWm9YcWJIWkdldyJ9.eyJhdWQiOiI2OGRmNjZhNC1jYWQ5LTRiZmQtODcyYi1jNmRkZGUwMGQ2YjIiLCJpc3MiOiJodHRwczovL2xvZ2luLm1pY3Jvc29mdG9ubGluZS5jb20vNzJmOTg4YmYtODZmMS00MWFmLTkxYWItMmQ3Y2QwMTFkYjQ3L3YyLjAiLCJpYXQiOjE2OTQ2ODI1OTYsIm5iZiI6MTY5NDY4MjU5NiwiZXhwIjoxNjk0Njg3MzY2LCJhaW8iOiJBYlFBUy84VUFBQUFGYXU4UXBRd3hBT243NUV4Nm83cnVnbWxXSkxkWVpaV1ZSKzM4VHdaUzg2NzBacWh2aHhDeDlZeGc2YVBoT1FsUSt0Z3JGMjAvNDIxeTBOVGFBTG9XUEVINzRCMmpRNXFWdVdPbUNKYTVMUzdPMjlKZXlCY1owaG9RSEREczUyc2U2S3RWV0xUWXNRZ0Z2ZDdFY090dW43TmpJaDlxL2owNzcvTi9QYXQzSXRwdURKZk9kcllacm0yTUZMRk5CREVUNnhaN3U2RjdwR00zYTAxdmZLSlZUSCtRNUdNWDJ2Ky9NR0owVFZjWmprPSIsImF6cCI6IjY4ZGY2NmE0LWNhZDktNGJmZC04NzJiLWM2ZGRkZTAwZDZiMiIsImF6cGFjciI6IjAiLCJlbWFpbCI6ImtoeWF0aW9zd2FsQG1pY3Jvc29mdC5jb20iLCJuYW1lIjoiS2h5YXRpIE9zd2FsIiwib2lkIjoiYWViYjcxNDMtYWQ2YS00N2Y2LWEyMTMtMmEwNDZhOTY2MmZhIiwicHJlZmVycmVkX3VzZXJuYW1lIjoia2h5YXRpb3N3YWxAbWljcm9zb2Z0LmNvbSIsInJoIjoiMC5BUUVBdjRqNWN2R0dyMEdScXkxODBCSGJSNlJtMzJqWnl2MUxoeXZHM2Q0QTFySWFBQncuIiwic2NwIjoiYWNjZXNzIiwic3ViIjoiRUxCQ3RWMVFSd2o0cmwycGlhTTRiWEg1M3JIMWZoaTFTTURIRm8takRfbyIsInRpZCI6IjcyZjk4OGJmLTg2ZjEtNDFhZi05MWFiLTJkN2NkMDExZGI0NyIsInV0aSI6Ilg3bHdxdlRva2syMHNud2RPNkZEQUEiLCJ2ZXIiOiIyLjAiLCJ2ZXJpZmllZF9wcmltYXJ5X2VtYWlsIjpbImtoeWF0aW9zd2FsQG1pY3Jvc29mdC5jb20iXX0.XPM1Lg92O4tAqv6WsomtNiu1GHBwxizVFf8FWGE-wCueITVjJ0s4jxiLdgrhb9DM9u1A9TxrFf7fgfuIEH3Wk3L7In7GR1godj-lqkvJISF7NmfoIaLXC96gnjg61tzH_vhW--2LmVE4bcJakHzkgnnZXhcZM6cHWz0uySSNjMR4Ho-QBW9cwfNqpY9mfT3kgvUdzUw9BWufop1EzQ2TyicvA_B0NtvjgVVKlNobRoQlhy2spYNDapZBEGQi4Qg5f2X2PqGJMJbbXEB9YGhF0LVrUA337SUpzDPLHY4AmC38IHzt6g2V5FSr6-O8OjAQV5SlZMdkMM0NBydyIxxrew";
        public class ModelPrompt
        {
            [JsonPropertyName("prompt")]
            public string Prompt
            {
                get;
                set;
            }

            [JsonPropertyName("max_tokens")]
            public int MaxTokens
            {
                get;
                set;
            }

            [JsonPropertyName("temperature")]
            public double Temperature
            {
                get;
                set;
            }

            [JsonPropertyName("top_p")]
            public int TopP
            {
                get;
                set;
            }

            [JsonPropertyName("n")]
            public int N
            {
                get;
                set;
            }

            [JsonPropertyName("stream")]
            public bool Stream
            {
                get;
                set;
            }

            [JsonPropertyName("logprobs")]
            public object? LogProbs
            {
                get;
                set;
            }

            [JsonPropertyName("stop")]
            public string? Stop
            {
                get;
                set;
            }
        };

        public class Choice
        {
            [JsonPropertyName("text")]
            public string? Text
            {
                get;
                set;
            }

            [JsonPropertyName("index")]
            public int Index
            {
                get;
                set;
            }

            [JsonPropertyName("logprobs")]
            public object? LogProbs
            {
                get;
                set;
            }

            [JsonPropertyName("finish_reason")]
            public string? FinishReason
            {
                get;
                set;
            }
        }

        public class LLMResponse
        {
            [JsonPropertyName("choices")]
            public List<Choice>? Choices
            {
                get;
                set;
            }
        }

        public class StreamResponse
        {
            [JsonPropertyName("id")]
            public string? Id
            {
                get;
                set;
            }

            [JsonPropertyName("object")]
            public string? Object
            {
                get;
                set;
            }

            [JsonPropertyName("created")]
            public int Created
            {
                get;
                set;
            }

            [JsonPropertyName("choices")]
            public List<Choice>? Choices
            {
                get;
                set;
            }

            [JsonPropertyName("model")]
            public string? Model
            {
                get;
                set;
            }
        }

        public async Task<string> GetToken()
        {
            var accounts = await app.GetAccountsAsync();
            AuthenticationResult? result = null;
            if (accounts.Any())
            {
                var chosen = accounts.First();
                result = await app.AcquireTokenSilent(SCOPES, chosen).ExecuteAsync();
            }
            if (result == null)
            {
                result = await app.AcquireTokenInteractive(SCOPES).ExecuteAsync();

                /*result = await app.AcquireTokenWithDeviceCode(SCOPES,
                    deviceCodeResult => {
                        // This will print the message on the console which tells the user where to go sign-in using
                        // a separate browser and the code to enter once they sign in.
                        // The AcquireTokenWithDeviceCode() method will poll the server after firing this
                        // device code callback to look for the successful login of the user via that browser.
                        // This background polling (whose interval and timeout data is also provided as fields in the
                        // deviceCodeCallback class) will occur until:
                        // * The user has successfully logged in via browser and entered the proper code
                        // * The timeout specified by the server for the lifetime of this code (typically ~15 minutes) has been reached
                        // * The developing application calls the Cancel() method on a CancellationToken sent into the method.
                        //   If this occurs, an OperationCanceledException will be thrown (see catch below for more details).
                        Console.WriteLine(deviceCodeResult.Message);
                        return Task.FromResult(0);
                    }).ExecuteAsync();*/
            }

            return (result.AccessToken);
        }

        public async Task<string> SendRequest(string modelType, string requestData, string llmToken)
        {
            var token = llmToken;

            //var token = await GetToken();
            var httpClient = new HttpClient();
            var request = new HttpRequestMessage(HttpMethod.Post, ENDPOINT);

            request.Content = new StringContent(requestData, Encoding.UTF8, "application/json");
            request.Headers.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", token);
            request.Headers.Add("X-ModelType", modelType);

            var httpResponse = await httpClient.SendAsync(request);

            return await httpResponse.Content.ReadAsStringAsync(); ;
        }

        public async Task<string> SendStreamRequest(string modelType, string requestData)
        {
            //var token = await GetToken();
            var token = "eyJ0eXAiOiJKV1QiLCJhbGciOiJSUzI1NiIsImtpZCI6Ii1LSTNROW5OUjdiUm9meG1lWm9YcWJIWkdldyJ9.eyJhdWQiOiI2OGRmNjZhNC1jYWQ5LTRiZmQtODcyYi1jNmRkZGUwMGQ2YjIiLCJpc3MiOiJodHRwczovL2xvZ2luLm1pY3Jvc29mdG9ubGluZS5jb20vNzJmOTg4YmYtODZmMS00MWFmLTkxYWItMmQ3Y2QwMTFkYjQ3L3YyLjAiLCJpYXQiOjE2NzY1NDcxMTUsIm5iZiI6MTY3NjU0NzExNSwiZXhwIjoxNjc2NTUxMzAwLCJhaW8iOiJBVlFBcS84VEFBQUExSkd5SFNPNzlVaVpzYjZkUkJVRzhvcG42Q0VpK3JNVUQ2Q1VTY0VteEsyRnRiaGQzK0pYcVZFWnQ5QUdZQmt3aEtuUndkTnNEZjlWdlE1NjV1L0k2NlNHamt5T0U2VkdrdmxQT1dsL2xDaz0iLCJhenAiOiI2OGRmNjZhNC1jYWQ5LTRiZmQtODcyYi1jNmRkZGUwMGQ2YjIiLCJhenBhY3IiOiIwIiwiZW1haWwiOiJ5cmFpemFkYUBtaWNyb3NvZnQuY29tIiwibmFtZSI6Illhc2ggVmFyZGhhbiBSYWl6YWRhIiwib2lkIjoiNDc2OWIyNDMtOWQyNi00ZTExLWFmOTUtZjIxMGQyOTA4OGU4IiwicHJlZmVycmVkX3VzZXJuYW1lIjoieXJhaXphZGFAbWljcm9zb2Z0LmNvbSIsInJoIjoiMC5BUm9BdjRqNWN2R0dyMEdScXkxODBCSGJSNlJtMzJqWnl2MUxoeXZHM2Q0QTFySWFBTVkuIiwic2NwIjoiYWNjZXNzIiwic3ViIjoiRXRodDY0MHBlQnBHcnNmaWx0VHlOUWFlNUxuOHNGcGJ0SGZXNG5wWEZCQSIsInRpZCI6IjcyZjk4OGJmLTg2ZjEtNDFhZi05MWFiLTJkN2NkMDExZGI0NyIsInV0aSI6IjZZR05RV09SaFV1N0I5dXlBT2hPQUEiLCJ2ZXIiOiIyLjAiLCJ2ZXJpZmllZF9wcmltYXJ5X2VtYWlsIjpbInlyYWl6YWRhQG1pY3Jvc29mdC5jb20iXX0.X1sCTKqSHxKdY9QwTGn7GtB8Qgqh259a1EEDMPil5A4DJQnBRdLpsn8lxvxTMqsavMenDnygr6DYhHmKZhkyb0Pom4EXqzsyxahi5PN40oFmYvEdua08jCckzgMeERgm3Pq1zhhSdi9oCFnNWlAJlI_9JoT7yjuEQ8Wo_fwLgTNROFYnAb6TmHhKdGmdXdAO-JDuAFdotsdVvp9CkYk3vAmMsXoFdzIxfh5gl7W_1t2adn1WGrL0_HZ79rL5y0Zx--bzhSxMFHPQUDZshHY7GhxiqJhi3IFx9PS3CRWPhs1W22Y6NEcBs3KtAeiTVXYORSfsf8eZ4Us7jM0BbN8Ijw";
            var httpClient = new HttpClient();
            var request = new HttpRequestMessage(HttpMethod.Post, ENDPOINT);
            request.Content = new StringContent(requestData, Encoding.UTF8, "application/json");
            request.Headers.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", token);
            request.Headers.Add("X-ModelType", modelType);

            var httpResponse = await httpClient.SendAsync(request, HttpCompletionOption.ResponseHeadersRead);

            var stream = await httpResponse.Content.ReadAsStreamAsync();
            TextReader textReader = new StreamReader(stream);

            string? line;
            while ((line = await textReader.ReadLineAsync()) != null)
            {
                if (line.StartsWith("data: "))
                {
                    var lineData = line.Substring(6);
                    if (string.Equals(lineData, "[DONE]"))
                    {
                        break;
                    }

                    var result = JsonSerializer.Deserialize<StreamResponse>(line.Substring(6));
                    return result.Choices[0].Text;
                }
            }

            return "Some error occured";
        }

        public async Task Main()
        {
            string requestData = JsonSerializer.Serialize(new ModelPrompt
            {
                Prompt = "Seattle is",
                MaxTokens = 50,
                Temperature = 1,
                TopP = 1,
                N = 5,
                Stream = false,
                LogProbs = null,
                Stop = "\n"
            });
            // get the model response
            // available models are:
            // text-davinci-001 (GPT-3)
            // text-davinci-002 (GPT-3.5)
            // text-davinci-003 (GPT-3.51)
            // text-chat-davinci-002 (ChatGPT)

            //var response = await SendRequest("text-davinci-002", requestData);
            var response = await SendRequest("text-davinci-002", requestData, "");
            Console.WriteLine(response);

            var streamRequestData = JsonSerializer.Serialize(new ModelPrompt
            {
                Prompt = "Instruction: Given an input question, respond with syntactically correct c++. Be creative but the c++ must be correct. \nInput: Create a function in c++ to remove duplicate strings in a std::vector<std::string>\n",
                MaxTokens = 500,
                Temperature = 0.6,
                TopP = 1,
                N = 1,
                Stream = true,
                LogProbs = null,
                Stop = "\r\n"
            });

            await SendStreamRequest("text-davinci-003", streamRequestData);

        }

        private string ParseLLMResponseToString(string response)
        {
            LLMResponse responseObject = JsonConvert.DeserializeObject<LLMResponse>(response);
            return responseObject.Choices[0].Text;
        }

        public async Task<string> GetSummarizedText(string text)
        {
            string requestData = JsonSerializer.Serialize(new ModelPrompt
            {
                Prompt = text,
                MaxTokens = 500,
                Temperature = 0.6,
                TopP = 1,
                N = 1,
                Stream = false,
                LogProbs = null,
                Stop = ""
            });

            var response = await SendRequest("text-davinci-003", requestData, _llmToken);

            //Parse the responseText well and present it neatly.

            var summaryText = ParseLLMResponseToString(response);

            return summaryText;
        }
    }


}
