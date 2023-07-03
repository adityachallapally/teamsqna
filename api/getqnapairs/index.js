
const { OpenAIClient, AzureKeyCredential } = require("@azure/openai");



module.exports = async function (context, req, config) {
    const res = {
        status: 200,
        body: {},
    };
    res.body = "only this part works";
    let qnapairs;

    const client = new OpenAIClient(
        "https://yammer-southus.openai.azure.com/",
        new AzureKeyCredential("1d635b5d42074364982685ae5afde51c")
    );

    const messages = [
        { role: "system", content: "You are a helpful assistant. You will talk like a pirate." },
        { role: "user", content: "Can you help me?" },
        { role: "assistant", content: "Arrrr! Of course, me hearty! What can I do for ye?" },
        { role: "user", content: "What's the best way to train a parrot?" },
      ];

    const deploymentId = "gpt-35-turbo";

    const events = await client.listChatCompletions(deploymentId, messages, { maxTokens: 128 });

    res.body = events;

    messsagestring = req.body.resultarray.join(",");

    //connect to azure openai sequence

    //call it with the string and put it in a try/catch

    //then parse the output string into an array of strings

    //return it to the client

    return res;

}