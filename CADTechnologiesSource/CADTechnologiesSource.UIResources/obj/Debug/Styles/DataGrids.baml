   M S B A M L   `   `   `  ????\  XCADTechnologiesSource.UIResources, Version=1.0.0.0, Culture=neutral, PublicKeyToken=nullc6clr-namespace:CADTechnologiesSource.UIResources.Styles(CADTechnologiesSource.UIResources.Styles  5      R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089  ?? MergedDictionaries ?? Source   
MMDataGrid ?? 
TargetType&    System.Windows.Controls.DataGrid ?? Property ?? Value*  $System.Windows.Controls.DataGridCell  FontSizeBase ?? 
TargetType ?? Padding J? SnapsToDevicePixels ??Name	 ? ColumnDefinitions
 ? RowDefinitions    SelectAllCommand ?? 
ResourceId ?? TypeInTargetAssembly7  1System.Windows.Controls.DataGridHeadersVisibility
   All ?? ConverterParameter!    HeadersVisibilityConverter ?? 	Converter ?? RelativeSource ?? AncestorTypeG  ASystem.Windows.Controls.Primitives.DataGridColumnHeadersPresenter  Name   Column ??Name ?? CanContentScroll ??Name ?? ScrollableHeight ? Maximum ?? Orientation* ?? #ComputedVerticalScrollBarVisibility ?? ViewportHeight ?? ViewportSize ? Value  ?? Mode! ?? ScrollableWidth," ?? %ComputedHorizontalScrollBarVisibility# ?? ViewportWidth$ ?? Triggers% ]? 
Conditions& ?? Property' ?? Value'  !System.Windows.VisualStateManager(   VisualStateGroups%  System.Windows.VisualStateGroup)   States   System.Windows.VisualState*   
Storyboard+ g? Duration, ?? 
TargetName- ?? TargetProperty  SemiDarkBackground	. ?? To  PrimaryBlue  DarkBackground/ ??Name0 ?? Color1 ?Name2 J? Opacity3 ? Points4 ?? Stretch=  7System.Windows.Controls.Primitives.DataGridColumnHeader  PrimaryBlueBrush  SegoeUI_SemiBold  FontSizeLarge5 P? Property6 P? Value ! DataGridHeaderBackgroundBrush &	 "HighlightedSemiDarkBackgroundBrush 
 SemiDarkBackgroundBrush:  4System.Windows.Controls.Primitives.DataGridRowHeader)	  #System.Windows.Controls.DataGridRow7 ?? Path  SegoeUI_Regular8 ??Name ! HighlightedSemiDarkBackground?
  9System.Windows.Controls.Primitives.SelectiveScrollingGrid?  9System.Windows.Controls.Primitives.DataGridCellsPresenter9 	  
ItemsPanelA  ;System.Windows.Controls.Primitives.DataGridDetailsPresenter: 	  DetailsVisibility$; 
  SelectiveScrollingOrientation;  5System.Windows.Controls.SelectiveScrollingOrientation<   Vertical#=    RowDetailsScrollingConverter
>   Row?? 5      D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    :x,http://schemas.microsoft.com/winfx/2006/xaml     5      @local6clr-namespace:CADTechnologiesSource.UIResources.Styles  5        5      ?? 5   
   $ Buttons.xaml??5	      ?? 5
   
   $ Colors.xaml??5      ?? 5   
   $ 
Fonts.xaml??5      ?? 5   
   $ FontAwesomeStrings.xaml??5      %!  &	        7  (?? 
    5?   	   $" DataGridSelectAllButtonStyle??"   )7  7  7  7  7  '  ?    7  7  7  7 	 7 
 7 	 7 
 7 	 7 	 '  ?    7 
 '	  ?    7  7  7  7  7  7  7  7  7  7  7  ?? 5      "   5   	   .?5   
   ??  ?@  	CellStyle 5      ?? 5      "  5      .?5      ??  ? ??$ 0q?5      ?? 5       ? ??$ 0 0 0 0q?5       ?? 5!       ? ??8   5#      5$      5&      ?? 5'   
    ? ??
 ???? 5)      ?? 5*   
    ? ??
 ?????5,      ?? 5-   
    ? ??
 ???? 5/      ?? 50   
    ? ??$ 0q?52      ?? 53   
     ?@  RowDetailsVisibilityMode$ VisibleWhenSelected=?55      ?? 56   
    ? ?? . 58      ?? 59   
    ?@??PanningMode$
 Both=?5;      ?? 5<   
    ?@??IsFlicksEnabled .  5>      ?? 5?   
    ? ?? 5A      ?? 5B      "   5C      .R?5D      ?? #??z??5E      #??z??5F      #??z??5G      # z??5H       . ._?5J      ?? $ DG_ScrollViewer??5K      |?.  ??5M      ?? 5N   "   " ??5O   %   .R?5P   &   ? 	 5Q   *   ?? 5R   .   $
??Auto ?5S   1   ?? 5T   .   $??* ?5U   1   ?? 5V   .   $
??Auto ?5W   1   
 5Y   *   ?? 5Z   .   $
??Auto ?5[   1   ?? 5\   .   $??* ?5]   1   ?? 5^   .   $
??Auto ?5_   1   .+?5a   *   ?? #??Z 5b   -   |?.  ??5d   -   C? *?? $" DataGridSelectAllButtonStyle??"   +y?5e   -   ?? *HeadersVisibility+# Z # Z  ?? "   ??5f   -   ?? *CellsPanelHorizontalOffset+ ?? "     5g   *   $! PART_ColumnHeadersPresenter??5h   -   $??1??5i   -   y?5j   -   ?? *HeadersVisibility+# Z # Z  ?? "   ?? 5k   *   $! PART_ScrollContentPresenter??5l   -   # z??5m   -   $??2??5n   -   $??1??5o   -   ?? 5p   *   $ PART_VerticalScrollBar??5q   -   $??2??5r   -   # z 5s   -   $ Vertical=?5t   -   $??1??5u   -   #y?z 5v   -   # z 5x   -    5w   -   ?? *VerticalOffset+$  OneWay=? ?? *TemplatedParent+5x   -   ? 5y   *   $??1??5z   -   $??2??5{   -   	 5|   .   ?? 5}   2   ??5~   5   ?? **(NonFrozenColumnsViewportHorizontalOffset+ ?? "   ?? 5   2   $??* ?5?   5   .+?5?   .   ?? $ PART_HorizontalScrollBar??5?   1   $??1??5?   1   # z! 5?   1   $ 
Horizontal=?5?   1   #y?z" 5?   1   # z# 5?   1    5?   1   ?? *HorizontalOffset+$  OneWay=? ?? *TemplatedParent+5?   1   5?   +   5?   '   5?   #   .??5?      ?? # z 5?   !   5?      5?      5?      5?      $ 5?   
   ]? 5?      % 5?      ?? 5?      & ?@??
IsGrouping' . 5?      ?? 5?      "& ?@%?IsVirtualizingWhenGrouping' .  5?      ./?5?      ??  ? ?? .  5?      5?      5?      ?? 5?      " ??5?   	   .?5?   
   ??  ? ??
 ???? 5?      ?? 5?   
    ? ?? 5?      ?? 5?      " ??5?      .R?5?      ? ( 5?        5?      CommonStates??.) 5?   "     
Normal??5?   %     5?   "   	MouseOver??.* 5?   &   ?? .??5?   *   ?? $+ 0D?5?   -   $, Border??5?   -   $- (Fill).(Color)?5?   -   8.   5?   -   5?   '   5?   #     5?   "   Pressed??.* 5?   &   ?? .??5?   *   ?? $+ 0D?5?   -   $, Border??5?   -   $- (Fill).(Color)?5?   -   8.  5?   -   5?   '   5?   #     5?   "   Disabled??.* 5?   &   ?? .??5?   *   ?? $+ 0D?5?   -   $, Border??5?   -   $- (Background).(Color)?5?   -   8.  5?   -   5?   '   5?   #   5?      .+?5?      ?? $/ Border??5?      $??0??5?       . ??5?      ?? 5?   "   80  5?   %   ??5?      ?? 5?   "   80  5?   %   5?      ? 5?      $1 Arrow??5?      $??Right=?5?      $??8,8,3,3q?5?      $2 1Q?5?      3 ?   ??? ??? ??? ??? $4 Uniform=?5?      $??Bottom=?5?      ??5?      ?? 5?   "   #0 ?  5?   %   5?      5?      5?      5?      5?      ?? 5?      "  5?   	   .?5?   
   ??  ? ??
 ???? 5?      ?? 5?   
    ? ??8   5     ?? 5  
    ? ??8  5     ?? 5  
    ? ??8  5     ?? 5  
    ? ??$ 0 0 0 0q?5
     ?? 5  
    ? ??$ 0 0 0 0q?5     $ 5  
   P? 5     5 ?@ SortDirection6 5     Z? . ?5     ??  ? ??# ?  5     ?? 5      ? ??
 ???? 5     5     ]? 5     % 5     ?? 5     & ?@J?IsMouseOver' . 5     ?? 5     & ?@ SortDirection' 5      Z? ./?5"     ??  ? ??8  5$     ?? 5%      ? ??8  5'     5(     ]? 5*     % 5+     ?? 5,     & ?@J?IsMouseOver' . 5.     ?? 5/     & ?@ SortDirection' 51     Z? ./?53     ??  ? ??8  55     ?? 56      ? ??8  58     59     P? 5:     5 ?@ SortDirection$6 	AscendingX?5<     . ?5=     ??  ? ??8  5?     5@     P? 5A     5 ?@ SortDirection$6 
DescendingX?5C     . ?5D     ??  ? ??8  5F     5G     5I     ?? 5L     "  5M  	   .?5N  
   ??  ? ??8   5P     ?? 5Q  
    ? ??$ 30??5S     ?? 5T  
    ? ??
 ???? 5V     ?? 5W  
    ? ??$ 0 0 0 0q?5Y     5Z     ?? 5]     " 	 5^  	   .?5_  
   ??  ? ?? 5a     ?? 5b     #0 ?  5c     5e     ?? 5f  
    ?@	 
IsSelected 5h     ?? $-7 'DataContext.SelectedOnOffConflictLayers??? 5i  
    ? ??8   5k     ?? 5l  
    ? ??$ 5q?5n     ?? 5o  
    ?@J?SnapsToDevicePixels . 5q     ?? 5r  
    ?@=?ErrorTemplate 5t     Z? ?? 5u  
    ?@	 ValidationErrorTemplate 5w     ?? 5x     .R?5y     ?? 
???  ??$??2,0,0,0q?5{     $??!??5|     $??Center=?5}     5~     5?     ?? 5?  
    ? ?? 5?     ?? 5?     " 	 5?     .R?5?     ?? $8 
DGR_Border??5?     #??z??5?     #??z??5?      . ??5?     ?? 5?     #0 ?  5?  !   ( 5?       5?     CommonStates??.) 5?  "     
Normal??5?  %     5?  "   Normal_AlternatingRow??.* 5?  &   ?? .??5?  *   ?? $+ 0D?5?  -   $, 
DGR_Border??5?  -   $- (Background).(Color)?5?  -   8.  5?  -   5?  '   5?  #     5?  "   Normal_Selected??.* 5?  &   ?? .??5?  *   ?? $+ 0D?5?  -   $, 
DGR_Border??5?  -   $- (Background).(Color)?5?  -   8.  5?  -   5?  '   5?  #     5?  "   Unfocused_Selected??.* 5?  &   ?? .??5?  *   ?? $+ 0D?5?  -   $, 
DGR_Border??5?  -   $- (Background).(Color)?5?  -   8.  5?  -   5?  '   5?  #     5?  "   Normal_Editing??.* 5?  &   ?? .??5?  *   ?? $+ 0D?5?  -   $, 
DGR_Border??5?  -   $- (Background).(Color)?5?  -   8.  5?  -   5?  '   5?  #     5?  "   MouseOver_Editing??.* 5?  &   ?? .??5?  *   ?? $+ 0D?5?  -   $, 
DGR_Border??5?  -   $- (Background).(Color)?5?  -   8.  5?  -   5?  '   5?  #     5?  "   MouseOver_Unfocused_Editing??.* 5?  &   ?? .??5?  *   ?? $+ 0D?5?  -   $, 
DGR_Border??5?  -   $- (Background).(Color)?5?  -   8.  5?  -   5?  '   5?  #     5?  "   Unfocused_Editing??.* 5?  &   ?? .??5?  *   ?? $+ 0D?5?  -   $, 
DGR_Border??5?  -   $- (Background).(Color)?5?  -   8.  5?  -   5?  '   5?  #     5?  "   	MouseOver??.* 5?  &   ?? .??5?  *   ?? $+ 0D?5?  -   $, 
DGR_Border??5?  -   $- (Background).(Color)?5?  -   8.  5?  -   5?  '   5?  #     5?  "   MouseOver_Selected??.* 5?  &   ?? .??5?  *   ?? $+ 0D?5?  -   $, 
DGR_Border??5?  -   $- (Background).(Color)?5?  -   8. 	 5?  -   5?  '   5   #     5  "    MouseOver_Unfocused_Selected??.* 5  &   ?? .??5  *   ?? $+ 0D?5  -   $, 
DGR_Border??5  -   $- (Background).(Color)?5  -   8. 
 5	  -   5
  '   5  #   5     ._?5     
  	 5     ?? 5  "   $
??Auto ?5  %   ?? 5  "   $??* ?5  %   
 5     ?? 5  "   $??* ?5  %   ?? 5  "   $
??Auto ?5  %   .+?5       $??1??5  !   #??z9 5  !   $??5q?5  !   # z 5   !     5!     $??1??5"  !   $??1??5#  !   #y?z: 5$  !   ; 5%  !   ?? *AreRowDetailsFrozen+# Z< # Z=  ?? "     5)     $??2??5*  !   $; Vertical=?5+  !   y?5,  !   ?? *HeadersVisibility+# Z> # Z  ?? "   50     51     52     54     55     57     