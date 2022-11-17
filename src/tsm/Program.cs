// See https://aka.ms/new-console-template for more information

using services;
using Terminal.Gui;

Application.Init();
Application.Top.Add(
    new MainView(Application.Top)
    {
        Width = Dim.Fill(0),
        Height = Dim.Fill(1)
    }
);
Application.Run(Application.Top);
Application.Shutdown();