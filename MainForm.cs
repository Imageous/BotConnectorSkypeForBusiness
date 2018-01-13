using Microsoft.Bot.Connector.Emulator;
using Microsoft.Lync.Model;
using Microsoft.Lync.Model.Conversation;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Windows.Forms;
using Newtonsoft.Json;
using ApiAiSDK;
using ApiAiSDK.Model;


namespace MrSamwiseBot
{
    public partial class MainForm : Form
    {
        private LyncClient _lyncClient;
        private ConversationManager _conversationManager;
        private readonly ConnectorEmulator oConnectorEmulator = new ConnectorEmulator();
        private ApiAi apiAi;

        //user apiai object mapping for multiple user testing
        Dictionary<string, ApiAi> apiAi_dict = new Dictionary<string, ApiAi>();

        string response=string.Empty;

        private ApiAi Get_apiai_object_from_username(string username)
        {
            var config = new AIConfiguration("YOUR_CLIENT_ACCESS_TOKEN", SupportedLanguage.English);
            
            ApiAi stored_apiai_object = null;
            if (apiAi_dict.ContainsKey(username))
            {
                //MessageBox.Show("found apiai object");
                stored_apiai_object = apiAi_dict[username];
            }
            else
            {
                //MessageBox.Show("create apiai object");
                apiAi = new ApiAi(config);
                apiAi_dict.Add(username, apiAi);
                stored_apiai_object = apiAi;
            }
            return stored_apiai_object;
        }

        private void test_function()
        {
            string username = "test user 1";
            ApiAi apiAi = Get_apiai_object_from_username(username);

            List<AIContext> aicontext = new List<AIContext>();
            AIContext user_info = new AIContext();
            user_info.Name = "skype_username=test";
            aicontext.Add(user_info);
            List<Entity> entity = new List<Entity>();
            var extras_info = new RequestExtras(aicontext, entity);
            var response = apiAi.TextRequest("I m cold, from someone says", extras_info);

            Newtonsoft.Json.Linq.JObject res_data = (Newtonsoft.Json.Linq.JObject)response.Result.Fulfillment.Messages[0];
            string ret_text = res_data["speech"].ToString();
            MessageBox.Show("for " + username + " : " + ret_text);


            username = "test user 2";
            apiAi = Get_apiai_object_from_username("user2");

            response = apiAi.TextRequest("yes", extras_info);
            res_data = (Newtonsoft.Json.Linq.JObject)response.Result.Fulfillment.Messages[0];
            ret_text = res_data["speech"].ToString();
            MessageBox.Show("for " + username + " : " + ret_text);


            username = "test user 1";
            apiAi = Get_apiai_object_from_username(username);

            response = apiAi.TextRequest("no", extras_info);
            res_data = (Newtonsoft.Json.Linq.JObject)response.Result.Fulfillment.Messages[0];
            ret_text = res_data["speech"].ToString();
            MessageBox.Show("for " + username + " : " + ret_text);

            username = "test user 1";
            apiAi = Get_apiai_object_from_username(username);
            response = apiAi.TextRequest("332", extras_info);
            res_data = (Newtonsoft.Json.Linq.JObject)response.Result.Fulfillment.Messages[0];
            ret_text = res_data["speech"].ToString();
            MessageBox.Show(ret_text);
        }
        

        public MainForm()
        {
            InitializeComponent();
            //test_function();      
            try
            {
                _lyncClient = LyncClient.GetClient();
                _conversationManager = _lyncClient.ConversationManager;
                _conversationManager.ConversationAdded += ConversationAdded;
            }
            catch (Exception ex)
            {
                LogMessage("LyncClient", ex.Message);
            }
            try
            {
                oConnectorEmulator.Port = Properties.Settings.Default.ConnectorEmulatorPort;
                oConnectorEmulator.StartServer();
            }
            catch
            {
                MessageBox.Show(string.Format("I can not start as Port:{0} is used by another application", Properties.Settings.Default.ConnectorEmulatorPort));
            }

        }


        private void ConversationAdded(object sender, ConversationManagerEventArgs e)
        {
            var conversation = e.Conversation;
            conversation.ParticipantAdded += ParticipantAdded;
        }

        private void ParticipantAdded(object sender, ParticipantCollectionChangedEventArgs e)
        {
            var participant = e.Participant;
            if (participant.IsSelf)
            {
                return;
            }

            var instantMessageModality =
                e.Participant.Modalities[ModalityTypes.InstantMessage] as InstantMessageModality;
            instantMessageModality.InstantMessageReceived += InstantMessageReceived;
        }

        private async void InstantMessageReceived(object sender, MessageSentEventArgs e)
        {
            var text = e.Text.Replace(Environment.NewLine, string.Empty);
            string myRemoteParticipantUri = (sender as InstantMessageModality).Endpoint.Uri.Replace("sip:", string.Empty);
            var username_string = (sender as InstantMessageModality).Participant.Contact.GetContactInformation(ContactInformationType.DisplayName); ;
            try
            {
                //LogMessage(myRemoteParticipantUri, text);
                string strResult = "";
                apiAi = Get_apiai_object_from_username(myRemoteParticipantUri);

                List<AIContext> aicontext = new List<AIContext>();
                AIContext user_info = new AIContext();
                user_info.Name = "skype_username=" + username_string;
                aicontext.Add(user_info);
                List<Entity> entity = new List<Entity>();

                var extras_info = new RequestExtras(aicontext, entity);

                var response = apiAi.TextRequest(text, extras_info);
                Newtonsoft.Json.Linq.JObject res_data = (Newtonsoft.Json.Linq.JObject)response.Result.Fulfillment.Messages[0];
                string ret_text = res_data["speech"].ToString();
                strResult = ret_text;

                (sender as InstantMessageModality).BeginSendMessage(strResult, null, null);
            }
            catch(Exception ex)
            {
                //LogMessage(myRemoteParticipantUri, ex.Message);
            }
        }

 
        void StartConversation(string myRemoteParticipantUri, string MSG)
        {
           
            foreach (var con in _conversationManager.Conversations)
            {
                if (con.Participants.Where(p => p.Contact.Uri == "sip:" + myRemoteParticipantUri).Count() > 0)
                {
                    if (con.Participants.Count == 2)
                    {
                        con.End();
                        break;
                    }
                }
            }

            Conversation _Conversation = _conversationManager.AddConversation();
            _Conversation.AddParticipant(_lyncClient.ContactManager.GetContactByUri(myRemoteParticipantUri));


            Dictionary<InstantMessageContentType, String> messages = new Dictionary<InstantMessageContentType, String>();
            messages.Add(InstantMessageContentType.PlainText, MSG);
            InstantMessageModality m = (InstantMessageModality)_Conversation.Modalities[ModalityTypes.InstantMessage];
            m.BeginSendMessage(messages, null, messages);
            LogMessage(myRemoteParticipantUri, MSG);
        }

        private static void LogMessage(string myRemoteParticipantUri, string MSG)
        {
            //MessageBox.Show(myRemoteParticipantUri + " : " + MSG);
        }
    }
}
