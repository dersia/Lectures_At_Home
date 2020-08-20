import { HubConnection, HubConnectionBuilder, HttpTransportType, LogLevel, MessageType } from '@aspnet/signalr'
import { Url } from 'url';

let signalRTokenEndpoint = "https://excelcf.azurewebsites.net/api/NegotiateSignalR?code=PREakYerAygMyKaI9l9nsHmWKdluF8N4sZFNDXXvazDryTxn/CCqkg==";
let cloudFancyEndpoint = "https://excelcf.azurewebsites.net/api/Fancy";
let cloudFancyAuth = "7by0/NvjPkmLDcG0K8oyhUYm0EEpUK4gqe0qnWEfhZy6bfkL47NtXg==";
let cloudAddEndpoint = "https://excelcf.azurewebsites.net/api/Add";
let cloudAddAuth = "MHf9DoteE5eUSgOnf1du8DaDIDRbnEgL/iY3X920nVj5xu9nWSQkWA==";

/**
 * Adds two numbers.
 * @customfunction
 * @param first First number
 * @param second Second number
 * @returns The sum of the two numbers.
 */
/* global clearInterval, console, setInterval */

export function add(first: number, second: number): number {
  return first + second;
}

/**
 * Adds two numbers using Azure Functions.
 * @customfunction CloudAdd
 * @param first First number
 * @param second Second number
 * @returns The sum of the two numbers.
 */

export async function cloud_add(first: number, second: number): Promise<number> {
  return await onAzure(cloudAddEndpoint, null, {first: first.toString(), second: second.toString()}, cloudAddAuth);
}

/**
 * Adds two numbers using Azure Functions.
 * @customfunction FancyCloudAlgo
 * @param num number
 * @returns Generate the Factorial of Given Number.
 */

export async function fancy_fact(num: number): Promise<number> {
  return await onAzure(cloudFancyEndpoint, null, {number: num.toString() }, cloudFancyAuth);
}

/**
 * Displays the current time once a second.
 * @customfunction
 * @param invocation Custom function handler
 */
export function clock(invocation: CustomFunctions.StreamingInvocation<string>): void {
  const timer = setInterval(() => {
    const time = currentTime();
    invocation.setResult(time);
  }, 1000);

  invocation.onCanceled = () => {
    clearInterval(timer);
  };
}

/**
 * Returns the current time.
 * @returns String with the current time formatted for the current locale.
 */
export function currentTime(): string {
  return new Date().toLocaleTimeString();
}

/**
 * Increments a value once a second.
 * @customfunction
 * @param incrementBy Amount to increment
 * @param invocation Custom function handler
 */
export function increment(incrementBy: number, invocation: CustomFunctions.StreamingInvocation<number>): void {
  let result = 0;
  const timer = setInterval(() => {
    result += incrementBy;
    invocation.setResult(result);
  }, 1000);

  invocation.onCanceled = () => {
    clearInterval(timer);
  };
}

/**
 * Writes a message to console.log().
 * @customfunction LOG
 * @param message String to write.
 * @returns String to write.
 */
export function logMessage(message: string): string {
  console.log(message);

  return message;
}

async function onAzure(url: string, body: object = null, parameters: Record<string, string> = null, auth: string = null): Promise<number> {
  let headers = new Headers();
  headers.set("Content-Type", "application/json");
  if(auth != null) {
    headers.set('x-functions-key', auth);
  }
  let fetchOptions: RequestInit = {
    method: body != null ? 'post' : 'get',
    mode: 'cors',
    cache: 'no-cache',
    redirect: 'follow',
    referrerPolicy: 'no-referrer',
    headers: headers,
    body: body != null ? JSON.stringify(body) : null
  };
  let requestUrl = new URL(url);
  if(parameters != null) {
    let searchParams = new URLSearchParams(parameters);
    requestUrl.search = searchParams.toString();
  }
  var response = await fetch(requestUrl.toString(), fetchOptions);
  return response.text().then(text => Number.parseInt(text));
}

/**
 * Displays the current time once a second.
 * @customfunction CONNECT_TO_SIGNALR
 * @param invocation Custom function handler
 */
export async function initSignalR(channel: string, invocation: CustomFunctions.StreamingInvocation<string>) {
  try {
    const res = await getSignalRInfo();
    if(typeof res === "string") {
      var response = JSON.parse(res);
      var options = {
        accessTokenFactory: () => response.accessToken
      }
      var connection: HubConnection = new HubConnectionBuilder().withUrl(response.url, options).build();
      connection.on(channel, (message:any) => {
        console.log(message);
        invocation.setResult("Message received: " + message);
      });
      invocation.onCanceled = async () => { await connection.stop(); console.log("disconnected") };
      await connection.start();
      console.log("connected");
    }  
  }
  catch (error) {
    console.error(error);
  }
}

async function getSignalRInfo() {
  try {
    const res = await fetch(signalRTokenEndpoint);
    return await res.text();
  }
  catch (error) {
    return console.log(error);
  }
}
