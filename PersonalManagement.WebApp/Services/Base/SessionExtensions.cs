using Microsoft.AspNetCore.Http;
using Newtonsoft.Json;

namespace PersonalManagement.WebApp.Services.Base
{
    public class SessionHelper
    {
        private readonly IHttpContextAccessor _httpContextAccessor;
        private ISession _session => _httpContextAccessor.HttpContext.Session;
        public SessionHelper(IHttpContextAccessor httpContextAccessor)
        {
            _httpContextAccessor = httpContextAccessor;
        }

        public void SetString(string key, string value)
        {
            _session.SetString(key, value);
        }

        public string GetString(string key)
        {
            return _session.GetString(key);
        }

        public void RemoveItem(string key)
        {
            _session.Remove(key);
        }

        public void SetObjectAsJson(string key, object value)
        {
            _session.SetString(key, JsonConvert.SerializeObject(value));
        }

        public T GetObjectFromJson<T>(string key)
        {
            var value = _session.GetString(key);

            return value == null ? default(T) : JsonConvert.DeserializeObject<T>(value);
        }
    }
}
