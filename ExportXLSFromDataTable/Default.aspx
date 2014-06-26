<%@ Page Title="主页" Language="C#" MasterPageFile="~/Site.Master" AutoEventWireup="true" CodeBehind="Default.aspx.cs" Inherits="ExportXLSFromDataTable._Default" %>

<asp:Content runat="server" ID="FeaturedContent" ContentPlaceHolderID="FeaturedContent">
    <section class="featured">
        <div class="content-wrapper">
            <hgroup class="title">
                <h1><%: Title %>.</h1>
                <h2>修改此模板以快速开始创建 ASP.NET 应用程序。</h2>
            </hgroup>
            <p>
                若要了解有关 ASP.NET 的详细信息，请访问 <a href="http://asp.net" title="ASP.NET Website">http://asp.net</a>。
                该页提供 <mark>视频、教程和示例</mark> 以帮助你充分利用 ASP.NET。
                如果你对 ASP.NET 有任何疑问，请访问
                <a href="http://forums.asp.net/18.aspx" title="ASP.NET Forum">我们的论坛</a>。
            </p>
        </div>
    </section>
</asp:Content>
<asp:Content runat="server" ID="BodyContent" ContentPlaceHolderID="MainContent">
    <asp:Literal ID="litTest" runat="server">

    </asp:Literal>
</asp:Content>
