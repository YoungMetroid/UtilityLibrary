using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;
using RestSharp;

namespace UtilityLibrary.CredentialAPI
{
	public class CredentialAPI
	{
		public static string uri = "http://15.60.43.31:8080/creds/get-credential/";
		public string username { get; set; }
		public string secret_key { get; set; }
		public string token { get; set; }

		public CredentialAPI()
		{

		}
		public CredentialAPI(string username, string secret_key, string token)
		{
			this.username = username;
			this.secret_key = secret_key;
			this.token = token;
		}
		public Identity GetCredentials()
		{
			var client = new RestClient();
			var request = new RestRequest(uri, Method.POST);
			request.RequestFormat = DataFormat.Json;
			string encoded = System.Convert.ToBase64String(System.Text.Encoding.GetEncoding("ISO-8859-1").GetBytes(username + ":" + secret_key));
			request.AddHeader("Content-type", "application/json");
			request.AddHeader("Authorization", "Basic " + encoded);
			request.AddJsonBody(new { Token = token });
			IRestResponse response = client.Execute(request);
			Identity credential = JsonConvert.DeserializeObject<Identity>(response.Content);
			return credential;
		}
	}

	public class Identity
	{
		public string CredentialName { get; set; }
		public string CredentialPwd { get; set; }
	}


}
