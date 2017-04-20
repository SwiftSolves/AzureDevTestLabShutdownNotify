#r "Newtonsoft.Json"
 
using System;
using System.Net;
using Newtonsoft.Json;
using System.Collections.Generic;
 
public static async Task<object> Run(HttpRequestMessage req, TraceWriter log)
{
    log.Info($"Webhook was triggered!");
 
    string jsonContent = await req.Content.ReadAsStringAsync();
    InNotification inNotification;
    try {
        inNotification = JsonConvert.DeserializeObject<InNotification>(jsonContent);
    }
    catch {
        return req.CreateResponse(HttpStatusCode.BadRequest, new {
            error = "Please pass object conforming to auto-shutdown notification schema in the input object"
        });
    }
 
    OutNotification content = new OutNotification {
        title = $"DevTest Labs {inNotification.eventType} Notification",
        text = $"The virtual machine {inNotification.vmName} in lab {inNotification.labName} with subscriptionId {inNotification.subscriptionId} is scheduled for {inNotification.eventType} in 15 minutes. Machine user is {inNotification.owner}. [Skip]({inNotification.skipUrl}), [Delay 1 hour]({inNotification.delayUrl60}) or [Delay 2 hours]({inNotification.delayUrl120})"
        //Potential Actions does not work without a registration
        //potentialAction = new List<PotentialAction>() {new PotentialAction{name = "Skip", target = new List<string>() {inNotification.skipUrl}}}
    };
 
    string response = await SendNotification.SendNotificationTeams(content, log);
    return req.CreateResponse(HttpStatusCode.OK, new {response = response});
}
 
private class InNotification 
{
    public string skipUrl { get; set; }
    public string delayUrl60 { get; set; }
    public string delayUrl120 { get; set; }
    public string vmName { get; set; }
    public string guid { get; set; }
    public string owner { get; set; }
    public string eventType { get; set; }
    public string text { get; set; }
    public string subscriptionId { get; set; }
    public string resourceGroupName { get; set; }
    public string labName { get; set; }
}
 
private class OutNotification
{
    public string title { get; set; }
    public string text { get; set; }
    public List<PotentialAction> potentialAction { get; set; }
}
 
private class PotentialAction
{
    public PotentialAction()
    {
        context = "http://schema.org";
        type = "ViewAction";
    }
    [JsonProperty(PropertyName = "@context")]
    public string context { get; set; }
    [JsonProperty(PropertyName = "@type")]
    public string type { get; set; }
    public string name { get; set; }
    public List<string> target { get; set; }
}
 
private static class SendNotification {
    static HttpClient httpClient;
    static string teamsWebhook;
    
    static SendNotification() {
        httpClient = new System.Net.Http.HttpClient();
        teamsWebhook = System.Configuration.ConfigurationManager.AppSettings["TeamsWebhook"];
        httpClient.DefaultRequestHeaders.Add("conent-type","application/json");
    }
 
    public static async Task<string> SendNotificationTeams(OutNotification body, TraceWriter log)
    {   
        var json = JsonConvert.SerializeObject(body);
        var content = new StringContent(json, System.Text.Encoding.UTF8, "application/json");
        HttpResponseMessage response = await httpClient.PostAsync(teamsWebhook, content);
        string responseMessage = string.Empty;
        if (response.IsSuccessStatusCode) {
            responseMessage = $"HTTP call suceeded";
        }
        else {
            responseMessage = $"HTTP call failed";
        }
        log.Info(responseMessage);
        return responseMessage;
    }
}
