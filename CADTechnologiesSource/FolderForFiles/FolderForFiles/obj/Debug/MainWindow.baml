   M S B A M L   `   `   `  ????I  EFolderForFiles, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null/clr-namespace:FolderForFilesFolderForFiles  5      E'clr-namespace:FolderForFiles.ViewModelsFolderForFiles.ViewModels  5	          FolderForFiles.MainWindow   5      R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    5      R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     5      @d2http://schemas.microsoft.com/expression/blend/2008     5      Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     5      .localclr-namespace:FolderForFiles     5      6vm'clr-namespace:FolderForFiles.ViewModels     5	      ???#1E1E1E  ? Title$  Folder For Files??5       ? WindowStartupLocation$ CenterScreen=?5       ? SizeToContent$ WidthAndHeight=?5      .??5      ? 3   -FolderForFiles.ViewModels.MainWindowViewModel   Instance ? DataContext# Z 5   	    ? RowDefinitions 5   
   ?? 5      ?? 5      .+?5   
   ?? $??0??6   .+?5      ??  ?? Orientation$ Vertical=?5      .+?5      ?? 
???????$??5q?5      $??Find:??5      ?? 5      $??5q?5       $
??Left=?5!      $	??100??5"      ??5      ?? *
FindText+ ?? Mode$ TwoWay=? ?? UpdateSourceTrigger$ PropertyChanged=?5"      ?? 5#      
???????$??5q?5%      $??Replace:??5&      ?? 5'      $??5q?5)      $
??Left=?5*      $	??100??5+      ??5(      ?? *ReplaceText+$ TwoWay=?$ PropertyChanged=?5+      ?? 5,      $	??100??5-      $
??Left=?5.      $??	5 10 5 10q?50      $??Find/Replace??51      ??5/      ?? *FindReplaceCommand+51      52      54      ?? 56      $??2??57   	   $??5q?58   	   .+?59   
   ?? $ 
Horizontal=?5:      .+?5;      ?? 
???????$??5q?5=      $??Select Folder:??5>      ?? 5?      $??5q?5A      $??Select path??5B      ??5@      ?? *FindPathCommand+5B      5C      ?? 5D   
   $??5q?5F      $
??Left=?5G      $	??600??5H      ??5E      ?? *
UserPath+$ TwoWay=?$ PropertyChanged=?5H      ?? 5I   
   $??5q?5K      $
??Left=?5L      $??Create Folders??5M      ??5J      ?? *MakeFoldersCommand+5M      5N      5P      5R      