using Newtonsoft.Json;

namespace CompareIntuneBaselineCompare.Models
{
    public class Policy
    {
        [JsonProperty("displayName", NullValueHandling = NullValueHandling.Ignore)]
        public string DisplayName { get; set; }

        [JsonProperty("description", NullValueHandling = NullValueHandling.Ignore)]
        public string Description { get; set; }

        [JsonProperty("roleScopeTagIds", NullValueHandling = NullValueHandling.Ignore)]
        public List<string> RoleScopeTagIds { get; set; }

        [JsonProperty("TemplateDisplayName", NullValueHandling = NullValueHandling.Ignore)]
        public string TemplateDisplayName { get; set; }

        [JsonProperty("TemplateId", NullValueHandling = NullValueHandling.Ignore)]
        public string TemplateId { get; set; }

        [JsonProperty("versionInfo", NullValueHandling = NullValueHandling.Ignore)]
        public string VersionInfo { get; set; }

        [JsonProperty("settingsDelta", NullValueHandling = NullValueHandling.Ignore)]
        public List<SettingsDeltum> SettingsDelta { get; set; }
    }

    public class SettingsDeltum
    {
        [JsonProperty("@odata.type", NullValueHandling = NullValueHandling.Ignore)]
        public string OdataType { get; set; }

        [JsonProperty("id", NullValueHandling = NullValueHandling.Ignore)]
        public string Id { get; set; }

        [JsonProperty("definitionId", NullValueHandling = NullValueHandling.Ignore)]
        public string DefinitionId { get; set; }

        [JsonProperty("valueJson", NullValueHandling = NullValueHandling.Ignore)]
        public string ValueJson { get; set; }

        [JsonProperty("value", NullValueHandling = NullValueHandling.Ignore)]
        public object Value { get; set; }
        [JsonProperty("implementationId", NullValueHandling = NullValueHandling.Ignore)]
        public object ImplementationId { get; set; }
    }


}
