   M S B A M L   `   `   `  ����H  DMahApps.Metro, Version=1.6.5.1, Culture=neutral, PublicKeyToken=nullV;clr-namespace:MahApps.Metro.Controls;assembly=MahApps.MetroMahApps.Metro.Controls  5   	   F BEasyWord.UI, Version=1.0.0.6, Culture=neutral, PublicKeyToken=null)clr-namespace:EasyWord.UIEasyWord.UI 5   	      EasyWord.UI.MainWindow   5      R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    5   	   R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     5   	   @d2http://schemas.microsoft.com/expression/blend/2008     5   	   Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     5   	   Mmetro;clr-namespace:MahApps.Metro.Controls;assembly=MahApps.Metro     5   	   +localclr-namespace:EasyWord.UI     5   	     � WindowStartupLocation$  CenterScreen=�5	       � 
ResizeMode$ NoResize=�6:      WinTitle � Title# [  5
   	   $	��450��6+   $	��600��68   .��5      � .+�5   
   �� -   5         ��Name$ 
btnReplace��5        ReplaceBtnTxt#��[ 6%   $
��Left=�6N   $��455,311,0,0q�5      $	��Top=�6&   $	��100��6>   $��25��6M   6\   �� 5   
   -   5        $ btnSelectFile_Replace��5        	SelBtnTxt#��[ 60   $
��Left=�6U   $��
28,311,0,0q�6p   $	��Top=�6�   $	��100��6�   $��25��6�   6�     EasyWord.UI.BusyIndicator  5   
   -   5          Name$ busyCtl��5      $��Center=�6/   $��Center=�6L   $	��100��6g   $��0,0,0,0q�6t    K� ZIndex$	 999��6�   $	��100��6�   �� 5   
   -   5        $��chkSelectDir��5      $.��(替换同目录下所有word/excel文档��6'   $
��Left=�6F   $��153,315,0,0q�6a   $	��Top=�6v   &   System.Windows.Controls.DataGrid  5   
   -   5        $��gridReplace��5   ?   
�������$��1q�6+      CanUserAddRows . $��Center=�6h      AutoGenerateColumns .  $	��194��5      $��
34,20,34,0q�6       GridLinesVisibility$	 All=�64   $	��Top=�6N   $	��526��6f   	   Columns	 5      4  .System.Windows.Controls.DataGridTemplateColumn  5      ,  &System.Windows.Controls.DataGridColumn
   Header$
 查找��6)      Width5  /System.Windows.Controls.DataGridLengthConverter$
 200* 65      CellTemplate 5      �� 5      .R�5      �� ��6(   ��  �� Path$
 From�6E   5         CellEditingTemplate 5      �� 5      .R�5      ��  �� AcceptsReturn . ��6&   �� $
 From� �� Mode$ TwoWay=� �� UpdateSourceTrigger$ PropertyChanged=�6�   5      5        5"      $
 	替换成��6)   $
 200* 66    5#      �� 5$      .R�5%      �� ��6(   �� $ To�6C   5&       5(      �� 5)      .R�5*      ��  . ��6&   �� $ To�$ TwoWay=�$ PropertyChanged=�6�   5+      5-      5/      �� 50   
   $��文件名替换：��6   $
��Left=�6!   $��
39,248,0,0q�6<   $	��Top=�6P   �� 51   
   -   5        $��txtFileFrom��51      $
��Left=�6%   $��23��6@   $��121,247,0,0q�6L    �� TextWrapping$
 Wrap=�6a   $�� ��6u   $	��Top=�6}   $	��120��6�   �� 52   
   -   5        $��	txtFileTo��52      $
��Left=�6#   $��23��6>   $��298,247,0,0q�6J   $
 Wrap=�6_   $�� ��6s   $	��Top=�6{   $	��120��6�   �� 53   
   $��	替换成��6   $
��Left=�6   $��247,248,0,0q�69   $	��Top=�6N   �� 54   
   -   5        $��btnAdd��54      $��	添加行��64   $
��Left=�6B   $��480,248,0,0q�6]   $	��Top=�6r   $��75��6�   �� 55   
   -	   5        $��tbMsg��55      $
��Left=�6!   $��
28,346,0,0q�6<   $
��Wrap=�6P   $��信息��6d   $	��Top=�6n   $	��527��6�   $��50��6�   57      58      