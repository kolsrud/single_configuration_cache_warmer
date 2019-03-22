using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;

namespace CacheWarmer
{
	public class Selection
	{
		public string Field;
		public List<string> Values;

		public Selection(string value)
		{
			var elements = WebUtility.UrlDecode(value).Split(',').ToArray();
			Field = elements.First();
			Values = elements.Skip(1).ToList();
		}
	}

	public struct SingleConfiguration
	{
		public readonly Uri Uri;
		public string Scheme => Uri.Scheme;
		public string Url => Uri.Host;
		public readonly string VirtualProxy;
		public readonly string AppId;
		public readonly string SheetId;
		public readonly List<Selection> Selections;

		public SingleConfiguration(Uri singleUri)
		{
			Uri = singleUri;
			VirtualProxy = singleUri.LocalPath.Substring(0, singleUri.LocalPath.IndexOf("single")).Trim('/');
			AppId = null;
			SheetId = null;
			Selections = new List<Selection>();
			foreach (var query in singleUri.Query.TrimStart('?').Split('&').Select(element => element.Split('=')))
			{
				if (query.Length != 2)
					throw new Exception("Unsupported empty query parameter in url: " + WebUtility.UrlDecode(singleUri.Query));

				var key = query[0];
				var value = query[1];
				switch (key)
				{
					case "appid":
						AppId = value;
						break;
					case "sheet":
						SheetId = value;
						break;
					case "select":
                        if (value != "clearall")
    						Selections.Add(new Selection(value));
						break;
				}
			}
		}
	}

}