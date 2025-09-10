namespace CAF.GstMatching.Web.Common
{
    public static class MySession
    {
        private static IHttpContextAccessor _httpContextAccessor;

        public static void Configure(IHttpContextAccessor httpContextAccessor)
        {
            _httpContextAccessor = httpContextAccessor;
        }

        public static CurrentSession Current => new CurrentSession(_httpContextAccessor?.HttpContext?.Session);

        public static string UserName
        {
            get => _httpContextAccessor?.HttpContext?.Session.GetString("UserName") ?? string.Empty;
            set
            {
                if (_httpContextAccessor?.HttpContext?.Session != null)
                {
                    _httpContextAccessor.HttpContext.Session.SetString("UserName", value);
                }
            }
        }
        public class CurrentSession
        {
            private readonly ISession _session;

            public CurrentSession(ISession session)
            {
                _session = session;
            }

            public string UserName
            {
                get => _session.GetString("UserName") ?? string.Empty;
                set => _session.SetString("UserName", value);
            } 

            public string Email
            {
                get => _session.GetString("Email") ?? string.Empty;
                set => _session.SetString("Email", value);
            }

            public string UserCode
            {
                get => _session.GetString("UserCode") ?? string.Empty;
                set => _session.SetString("UserCode", value);
            }

            public string Taskseqid
            {
                get => _session.GetString("Taskseqid") ?? string.Empty;
                set => _session.SetString("Taskseqid", value);
            }

            public int? age
            {
                get => _session.GetInt32("Age");
                set => _session.SetInt32("Age", value ?? 0);
            }

            public string Level1
            {
                get => _session.GetString("Level1") ?? string.Empty;
                set => _session.SetString("Level1", value);
            }

            public string Level2
            {
                get => _session.GetString("Level2") ?? string.Empty;
                set => _session.SetString("Level2", value);
            }
            public string gstin
            {
                get => _session.GetString("gstin") ?? string.Empty;
                set => _session.SetString("gstin", value);
            }
            public string loginpassword
            {
                get => _session.GetString("password") ?? string.Empty;
                set => _session.SetString("password", value);
            }
            public string passwordChanged
            {
                get => _session.GetString("passwordChanged") ?? string.Empty;
                set => _session.SetString("passwordChanged", value);
            }
        }
    }
}