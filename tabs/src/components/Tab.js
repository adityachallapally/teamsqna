// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import React from "react";
import "./App.css";
import "./Tab.css";
import { createMicrosoftGraphClientWithCredential } from "@microsoft/teamsfx";
import {
  BearerTokenAuthProvider,
  createApiClient,
  TeamsUserCredential,
} from "@microsoft/teamsfx";
import Profile from "./Profile";
import Creator from "./Creator";
import {
  Checkbox,
  Button,
  Input,
  Menu,
  MenuTrigger,
  MenuButton,
  MenuPopover,
  MenuList,
  MenuItem,
} from "@fluentui/react-components";
import { Notepad20Regular, ContactCard20Regular } from "@fluentui/react-icons";
import noItemimage from "../images/no-item.png";
import { app } from "@microsoft/teams-js";
import config from "./lib/config";

import { OpenAIClient, AzureKeyCredential } from "@azure/openai";

// Now you can use OpenAIClient and AzureKeyCredential in your code

//const { OpenAIClient, AzureKeyCredential } = require("@azure/openai");


class Tab extends React.Component {
  constructor(props) {
    super(props);
    this.state = {
      userInfo: {},
      items: [],
      newItemDescription: "",
      showLoginPage: undefined,
      isAddingItem: false,
      initialized: false,
    };
  }

  async componentDidMount() {
    await this.initTeamsFx();
    await this.initData();
  }

  async initTeamsFx() {
    const authConfig = {
      clientId: config.clientId,
      initiateLoginEndpoint: config.initiateLoginEndpoint,
    };

    const credential = new TeamsUserCredential(authConfig);
    const userInfo = await credential.getUserInfo();

    this.setState({
      userInfo: userInfo,
    });

    this.scope = ["User.Read", "User.ReadBasic.All"];
    this.channelOrChatId = await this.getChannelOrChatId();
    this.credential = credential;

    const apiBaseUrl = config.apiEndpoint + "/api/";
    // createApiClient(...) creates an Axios instance which uses BearerTokenAuthProvider to inject token to request header
    const apiClient = createApiClient(
      apiBaseUrl,
      new BearerTokenAuthProvider(
        async () => (await credential.getToken("")).token
      )
    );
    this.apiClient = apiClient;
  }

  async initData() {
    if (!(await this.checkIsConsentNeeded())) {
      await this.getItems();
    }
  }

  async loginBtnClick() {
    try {
      // Popup login page to get user's access token
      await this.credential.login(this.scope);
    } catch (err) {
      if (err instanceof Error && err.message?.includes("CancelledByUser")) {
        const helpLink = "https://aka.ms/teamsfx-auth-code-flow";
        err.message +=
          '\nIf you see "AADSTS50011: The reply URL specified in the request does not match the reply URLs configured for the application" ' +
          "in the popup window, you may be using unmatched version for TeamsFx SDK (version >= 0.5.0) and Teams Toolkit (version < 3.3.0) or " +
          `cli (version < 0.11.0). Please refer to the help link for how to fix the issue: ${helpLink}`;
      }

      alert("Login failed: " + err);
      return;
    }
    await this.initData();
  }

  async checkIsConsentNeeded() {
    try {
      await this.credential.getToken(this.scope);
    } catch (error) {
      this.setState({
        showLoginPage: true,
      });
      return true;
    }
    this.setState({
      showLoginPage: false,
    });
    return false;
  }

  async getChannelOrChatId() {
    return new Promise((resolve) => {
      app.getContext().then((context) => {
        if (context.channelId) {
          resolve(context.channelId);
        } else if (context.channel?.id) {
          resolve(context.channel.id);
        } else if (context.chatId) {
          resolve(context.chatId);
        } else if (context.chat?.id) {
          resolve(context.chat.id);
        } else {
          resolve(this.state.userInfo.objectId);
        }
      });
    });
  }

  async callFunctionWithErrorHandling(command, method, options, params) {
    var message = [];
    var funcErrorMsg = "";
    try {
      const response = await this.apiClient.request({
        method: method,
        url: command,
        data: options,
        params,
      });
      message = response.data;
    } catch (err) {
      if (err.response && err.response.status && err.response.status === 404) {
        funcErrorMsg =
          'There may be a problem with the deployment of Azure Function App, please deploy Azure Function (Run command palette "TeamsFx - Deploy Package") first before running this App';
      } else if (err.message === "Network Error") {
        funcErrorMsg =
          "Cannot call Azure Function due to network error, please check your network connection status and ";
        if (err.config.url.indexOf("localhost") >= 0) {
          funcErrorMsg +=
            'make sure to start Azure Function locally (Run "npm run start" command inside api folder from terminal) first before running this App';
        } else {
          funcErrorMsg +=
            'make sure to provision and deploy Azure Function (Run command palette "TeamsFx - Provision Resource" and "TeamsFx - Deploy Package") first before running this App';
        }
      } else {
        funcErrorMsg = err.toString();
        if (err.response?.data?.error) {
          funcErrorMsg += ": " + err.response.data.error;
        }
        alert(funcErrorMsg);
      }
    }
    return message;
  }

  async getItems() {
    // Use client TeamsFx SDK to call "todo" Azure Function in "get" method to get all todo list which belong to user oid
    let result = await this.callFunctionWithErrorHandling(
      "todo",
      "get",
      undefined,
      { channelOrChatId: this.channelOrChatId }
    );
    if ("Error" === result) {
      throw new Error(
        "todo Function failed, please check Azure Functions log for details!"
      );
    } else {
      this.setState({
        items: result,
        initialized: true,
      });
    }
  }

  async onAddItem() {
    const newItems = JSON.parse(JSON.stringify(this.state.items));
    newItems.push({
      description: this.state.newItemDescription,
      objectId: this.state.userInfo.objectId,
    });
    this.setState({
      newItemDescription: "",
      items: newItems,
    });

    // Use client TeamsFx SDK to call "todo" Azure Function in "post" method to insert a new todo item under user oid
    await this.callFunctionWithErrorHandling("todo", "post", {
      description: this.state.newItemDescription,
      isCompleted: false,
      channelOrChatId: this.channelOrChatId,
    });
    this.refresh();
  }

  async onUpdateItem(id, description) {
    // Use client TeamsFx SDK to call "todo" Azure Function in "put" method to update a todo item
    await this.callFunctionWithErrorHandling("todo", "put", {
      id,
      description,
    });
  }

  async onDeleteItem(id) {
    const newItems = this.state.items.filter((item) => item.id !== id);
    this.setState({
      items: newItems,
    });
    // Use client TeamsFx SDK to call "todo" Azure Function in "delete" method to delete a todo item
    await this.callFunctionWithErrorHandling("todo", "delete", { id });
  }

  async onCompletionStatusChange(id, index, isCompleted) {
    this.handleInputChange(index, "isCompleted", isCompleted);
    // Use client TeamsFx SDK to call "todo" Azure Function in "put" method to update a todo item to completed
    await this.callFunctionWithErrorHandling("todo", "put", {
      id,
      isCompleted,
    });
  }

  handleInputChange(index, property, value) {
    const newItems = JSON.parse(JSON.stringify(this.state.items));
    newItems[index][property] = value;
    this.setState({
      items: newItems,
    });
  }

  handleClick = async () => {
    //get messages from the channel from the user profile
    //get messages from the channel from the user profile
    let result = await this.callFunctionWithErrorHandling("getuserprofile", "post", {});
    console.log(result.join("."));

    //get the Q&A pairs
    const extractedPairs = await this.qnaextraction();
    /*
    const data = {
      "0": {
        "Q": "What is the weather today?",
        "A": "It's warm."
      },
      "1": {
        "Q": "What did you add at the top of this channel?",
        "A": "I added a tab at the top of this channel. Check it out!"
      }
    };
    
    const extractedPairs = Object.keys(data).map((key) => data[key]);
    */
    console.log(extractedPairs);

    for (const pair of extractedPairs) {
      const requestBodyQ = {
        description: pair.Q,
        isCompleted: false,
        channelOrChatId: this.channelOrChatId,
      };

      await this.callFunctionWithErrorHandling("todo", "post", requestBodyQ);
      console.log(pair.Q);

      const requestBodyA = {
        description: pair.A,
        isCompleted: false,
        channelOrChatId: this.channelOrChatId,
      };

      await this.callFunctionWithErrorHandling("todo", "post", requestBodyA);
      console.log(pair.A);

      console.log("-------------------------");
    }

    this.refresh();

    // let qnapairs = await this.callFunctionWithErrorHandling("getqnapairs", "post", {
    //   resultarray: result,});

    // console.log(qnapairs);

    //put the Q&A pairs into a db
  }


  qnaextraction = async () => {

    const client = new OpenAIClient(
      "https://yammer-southus.openai.azure.com/",
      new AzureKeyCredential("1d635b5d42074364982685ae5afde51c")
    );

    const textToSummarize = `
  I added a tab at the top of this channel. Check it out!<attachment id="tab::b2f82ee4-7308-4bc1-948a-1023469e9dab"></attachment>.I added a tab at the top of this channel. Check it out!<attachment id="tab::cf0ba183-a8e5-4911-9a82-5cca269537fe"></attachment>first replysecond replythird reply,,,.I added a tab at the top of this channel. Check it out!<attachment id="tab::1dafcda6-c02a-4707-92d4-303bd4f23ec3"></attachment>.I added a tab at the top of this channel. Check it out!<attachment id="tab::57cb0845-8aad-4389-bb69-5e822aa852a3"></attachment>.I added a tab at the top of this channel. Check it out!<attachment id="tab::cbb4fc4c-a4f6-4d04-8e49-10b5f7ace4ff"></attachment>.I added a tab at the top of this channel. Check it out!<attachment id="tab::6bef4315-97ee-47d8-b196-5d448506f2e0"></attachment>.what is the weather today?it's warm,
:`;

    const summarizationPrompt = [`
If there are no questions in the text, just return an empty JSON object. Create Q&A pairs from the following text, with this JSON output:
'{
  "0":
  {
      "Q": question text,
      "A": answer text
  },
  "1":
  {
      "Q": question text,
      "A": answer text
  }
}'
Do not include anything else in the output. Only include the response in that JSON format and nothing else. Don't describe the text, just provide only what's stated here. If there are no questions in the text, just return an empty JSON object. Always return at least a question.
Text:
""""""
${textToSummarize}
""""""

Summary:
`];

    //console.log(`Input: ${summarizationPrompt}`);

    const deploymentName = "text-davinci-003";

    const { choices } = await client.getCompletions(deploymentName, summarizationPrompt, {
      maxTokens: 1000
    });
    const completion = choices[0].text;
    console.log(completion);
    const json_completion = JSON.parse(completion);
    console.log(json_completion);

    const pairsArray = [];
    for (const key in json_completion) {
      let value = json_completion[key].A;
      let replacedAnswer = value.includes("attachment") || value.includes("'") ? "There is an attachment" : value;
      console.log(replacedAnswer);

      let question = json_completion[key].Q;
      let replacedQuestion = question.includes("attachment") || question.includes("'") ? "There is an attachment" : question;
      console.log(replacedQuestion);

      const pair = {
        Q: replacedQuestion,
        A: replacedAnswer
      };
      pairsArray.push(pair);
    }

    return pairsArray;
  }

  async refresh() {
    await this.getItems();
  }

  render() {
    const items = this.state.items?.map((item, index) => (
      <div key={index} className="item">
        <div className="description">
          <Input
            value={this.state.items[index].description}
            onChange={(e) =>
              this.handleInputChange(index, "description", e.target.value)
            }
            onKeyDown={(e) => {
              if (e.key === "Enter") {
                this.onUpdateItem(item.id, this.state.items[index].description);
                e.target.blur();
              }
            }}
            onBlur={() =>
              this.onUpdateItem(item.id, this.state.items[index].description)
            }
            className={
              "text" +
              (this.state.items[index].isCompleted ? " is-completed" : "")
            }
          />
        </div>
        <Creator
          objectId={item.objectId}
          teamsUserCredential={this.credential}
          scope={this.scope}
        />
        <div className="action">
          <Menu>
            <MenuTrigger disableButtonEnhancement>
              <MenuButton appearance="subtle" menuIcon={null}>
                ...
              </MenuButton>
            </MenuTrigger>
            <MenuPopover>
              <MenuList>
                <MenuItem onClick={() => this.onDeleteItem(item.id)}>
                  Delete
                </MenuItem>
              </MenuList>
            </MenuPopover>
          </Menu>
        </div>
      </div>
    ));

    return (
      <div>
        {this.state.showLoginPage === false && (
          <div className="flex-container">
            <div className="todo-col">
              <div className="todo">
                <div className="header">
                  <div className="title">
                    <h2>FAQ</h2>
                    <span>
                      {this.state.items.length} item
                      {this.state.items.length === 1 ? "" : "s"}
                    </span>
                  </div>

                  <div className="import-button">
                    <Button
                      appearance="primary"
                      onClick={this.handleClick}
                    >
                      + Import
                    </Button>
                  </div>

                  <div className="add-button">
                    <Button
                      appearance="primary"
                      onClick={() => this.setState({ isAddingItem: true })}
                    >
                      + Add task3
                    </Button>
                  </div>
                </div>

                {this.state.items.length > 0 && (
                  <div className="header-container">
                    <div className="note">
                      <Notepad20Regular />
                      <span>Q&A Pair</span>
                    </div>
                    <div className="created-by">
                      <ContactCard20Regular />
                      <span>Created By</span>
                    </div>
                  </div>
                )}

                {items}

                {this.state.isAddingItem && (
                  <div className="item add">
                    <div className="complete">
                      <Checkbox disabled className="is-completed-input" />
                    </div>
                    <div className="description">
                      <Input
                        autoFocus
                        type="text"
                        value={this.state.newItemDescription}
                        onChange={(e) =>
                          this.setState({ newItemDescription: e.target.value })
                        }
                        onKeyDown={(e) => {
                          if (e.key === "Enter") {
                            this.onAddItem();
                          }
                        }}
                        onBlur={() => {
                          if (this.state.newItemDescription) {
                            this.onAddItem();
                          }
                          this.setState({
                            isAddingItem: false,
                          });
                        }}
                        className="text"
                      />
                    </div>
                  </div>
                )}

                {this.state.initialized &&
                  !this.state.items.length &&
                  !this.state.isAddingItem && (
                    <div className="no-item">
                      <div>
                        <img src={noItemimage} alt="no item" />
                      </div>
                      <div>
                        <h2>No tasks</h2>
                        <p>Add more tasks to make you day productive.</p>
                      </div>
                    </div>
                  )}
              </div>
            </div>
          </div>
        )}

        {this.state.showLoginPage === true && (
          <div className="auth">
            <Profile userInfo={this.state.userInfo} />
            <h2>Welcome to To Do List App!</h2>
            <Button appearance="primary" onClick={() => this.loginBtnClick()}>
              Start
            </Button>
          </div>
        )}
      </div>
    );
  }
}
export default Tab;
