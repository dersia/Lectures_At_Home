import { HubConnection, HubConnectionBuilder, HttpTransportType, LogLevel, MessageType } from '@aspnet/signalr'
import { Url } from 'url';
import {v4 as uuid } from 'uuid';

const signalRTokenEndpoint:string = "https://excelcf.azurewebsites.net/api/NegotiateSignalR?code=PREakYerAygMyKaI9l9nsHmWKdluF8N4sZFNDXXvazDryTxn/CCqkg==";
const cloudFancyEndpoint:string = "https://excelcf.azurewebsites.net/api/Fancy";
const cloudFancyAuth:string = "7by0/NvjPkmLDcG0K8oyhUYm0EEpUK4gqe0qnWEfhZy6bfkL47NtXg==";
const cloudAddEndpoint:string = "https://excelcf.azurewebsites.net/api/Add";
const cloudAddAuth:string = "MHf9DoteE5eUSgOnf1du8DaDIDRbnEgL/iY3X920nVj5xu9nWSQkWA==";
const apikey:string = "a0e2fa30075c4eac870f368f7316e5e3";
const endpoint:string = "https://api.cognitive.microsofttranslator.com/";

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

async function onAzure(url: string, body?: object, parameters?: Record<string, string>, auth?: string): Promise<number> {
  let headers = new Headers();
  headers.set("Content-Type", "application/json");
  if(auth !== null) {
    headers.set('x-functions-key', auth);
  }
  let fetchOptions: RequestInit = {
    method: body !== null ? 'post' : 'get',
    mode: 'cors',
    cache: 'no-cache',
    redirect: 'follow',
    referrerPolicy: 'no-referrer',
    headers: headers,
    body: body !== null ? JSON.stringify(body) : null
  };
  let requestUrl = new URL(url);
  if(parameters !== null) {
    let searchParams = new URLSearchParams(parameters);
    requestUrl.search = searchParams.toString();
  }
  var response = await fetch(requestUrl.toString(), fetchOptions);
  return response.text().then(text => Number.parseInt(text));
}

/**
 * Displays the current time once a second.
 * @customfunction CONNECT_TO_SIGNALR
 * @param channel channel to connect to
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

/**
 * Translate a Term using Microsoft Translate.
 * @customfunction translate
 * @param term Term to be translated
 * @param from original language of the Term
 * @param to language to which the Term is translated to
 * @returns Translated Term
 */

export async function translate(term: string, from: string, to: string): Promise<string> {
  const url = endpoint + "translate?api-version=3.0&to=" + to + "&from=" + from;
  const body = JSON.stringify([{ text: term }]);
  try {
    const res = await fetch(url, {
      method: 'POST',
      body: body,
      headers: {
        "Content-Type": "application/json",
        "Ocp-Apim-Subscription-Key": apikey,
        "X-ClientTraceId": uuid().toString(),
        "Ocp-Apim-Subscription-Region": "westeurope"
      }
    });
    const data = await res.json();
    console.log(data);
    if(data.error) {
      throw data.error.message;
    }
    return data[0].translations[0].text;
  } catch(error) {
    console.log(error);
    let errorT = typeof(error);
    throw error;
    return "error";
  }
}

/**
 * Translate a Term using Microsoft Translate (auto detect from).
 * @customfunction translate_AUTO
 * @param term Term to be translated
 * @param to language to which the Term is translated to
 */

export async function translateAuto(term: string, to: string): Promise<string> {
  const url = endpoint + "translate?api-version=3.0&to=" + to ;
  const body = JSON.stringify([{ text: term }]);
  try {
    const res = await fetch(url, {
      method: 'POST',
      body: body,
      headers: {
        "Content-Type": "application/json",
        "Ocp-Apim-Subscription-Key": apikey,
        "X-ClientTraceId": uuid().toString(),
        "Ocp-Apim-Subscription-Region": "westeurope"
      }
    });
    const data = await res.json();
    console.log(data);
    if(data.error) {
      return data.error.message;
    }
    return data[0].translations[0].text;
  } catch(error) {
    console.log(error);
    let errorT = typeof(error);
    throw error;
  }
}
