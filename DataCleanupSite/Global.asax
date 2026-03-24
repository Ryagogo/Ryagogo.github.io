<%@ Application Language="C#" %>
<%@ Import Namespace="System.Web.Routing" %>

<script runat="server">

    void Application_Start(object sender, EventArgs e)
    {
        // Application startup logic
    }

    void Application_End(object sender, EventArgs e)
    {
        // Application shutdown logic
    }

    void Application_Error(object sender, EventArgs e)
    {
        Exception ex = Server.GetLastError();
        if (ex != null)
        {
            // Log error details (extend to write to Event Log or file as needed)
            System.Diagnostics.Debug.WriteLine("[APP ERROR] " + ex.ToString());
        }
    }

    void Session_Start(object sender, EventArgs e) { }
    void Session_End(object sender, EventArgs e) { }

</script>
