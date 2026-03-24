<%@ Page Title="Error — DataCleanup Tool" Language="C#" MasterPageFile="~/Site.Master" AutoEventWireup="true" %>

<asp:Content ID="MainContent" ContentPlaceHolderID="MainContent" runat="server">
    <div class="wrapper" style="text-align:center;padding-top:80px;">
        <div style="font-size:64px;margin-bottom:24px;">⚠</div>
        <h1 style="font-family:var(--mono);font-size:28px;color:var(--danger);margin-bottom:12px;">Something went wrong</h1>
        <p style="color:var(--text-dim);font-size:15px;margin-bottom:32px;">
            An unexpected error occurred. Please try again or return to the home page.
        </p>
        <a href="~/Default.aspx" runat="server" class="btn btn-primary">← Back to Home</a>
    </div>
</asp:Content>
