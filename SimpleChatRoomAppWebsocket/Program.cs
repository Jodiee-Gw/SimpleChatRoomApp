using System;
using System.Net;
using System.Net.WebSockets;
using System.Text;

var builder = WebApplication.CreateBuilder(args);
var app = builder.Build();


app.UseWebSockets();

app.MapGet("/ws", async context =>
{
    if (context.WebSockets.IsWebSocketRequest)
    {
        using (var webSocket = await context.WebSockets.AcceptWebSocketAsync())
        {
            while (true)
            {
                var message = $"The current time : {DateTime.Now.ToString("HH:mm:ss")}";
                var bytes = Encoding.UTF8.GetBytes(message);
                var arraySegment = new ArraySegment<byte>(bytes, 0, bytes.Length);

                if (webSocket.State == WebSocketState.Open)
                {
                    await webSocket.SendAsync(arraySegment, WebSocketMessageType.Text, true, CancellationToken.None);
                }
                else if (webSocket.State == WebSocketState.Closed || webSocket.State == WebSocketState.Aborted)
                {
                    break;
                }
                Thread.Sleep(1000);
            }
            
        }
    }
    else
    {
        context.Response.StatusCode = (int)HttpStatusCode.BadRequest; 
    }

});

app.Run();
