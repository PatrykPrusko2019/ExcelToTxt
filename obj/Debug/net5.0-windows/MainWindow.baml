   M S B A M L   `   `   `  ����E  AExcelToTxt, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null'clr-namespace:ExcelToTxt
ExcelToTxt  5   	       ExcelToTxt.MainWindow   5   	   W SPresentationCore, Version=5.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=5.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35e aSystem.Windows.Controls.Ribbon, Version=5.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089R NWindowsBase, Version=5.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35F 9http://schemas.microsoft.com/winfx/2006/xaml/presentation     5   	   R NSystem.Xaml, Version=5.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     5   	   @d2http://schemas.microsoft.com/expression/blend/2008     5   	   Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     5   	   *localclr-namespace:ExcelToTxt     5   	     � Title$+  %Application convert excel to txt file��5   	   $	��250��67   $	��420��6D   .��5	      � .+�5   
   �� $��1. Choose excel file��6   $
��Left=�6/   $��	22,23,0,0q�6J   $	��Top=�6]   $	��111��6u   �� 5   
   -   5        $��	Open file��5      $
��Left=�6%   $��
174,26,0,0q�6@   $	��Top=�6T   6l   �� 5   
   -   5         ��Name$ image_ok_open_file��5      $
��Left=�6,   $��19��6G   $��
231,26,0,0q�6S   $	��Top=�6g   $��26��6   $��/img/image.jpg��6�    �� StretchDirection$
 Both=�6�   $y�Hidden=�6�   �� 5   
   $ ��2. Cylinders.xlsx / xls ->��6   $
��Left=�65   $��	22,54,0,0q�6P   $	��Top=�6c   $	��131��6{   �� 5   
   -   5        $��Convert cylinder��5      $
��Left=�6,   $��
291,57,0,0q�6G   $	��Top=�6[   6s   �� 5   
   -   5         ��Name$ name_Sheet_Cylinder��5      $
��Left=�6/   $��
174,58,0,0q�6J    �� TextWrapping$
 Wrap=�6^   $	��Top=�6r   $	��104��6�   4  .System.Windows.Automation.AutomationProperties   HelpText$ fill which spreadsheet��6�      Name$
 hehe��6�   �� 5   
   -   5        $ image_ok_cylinders��5      $
��Left=�6,   $��19��6G   $��
388,58,0,0q�6S   $	��Top=�6g   $��26��6   $��/img/image.jpg��6�   $
 Both=�6�   $y�Hidden=�6�   �� 5   
   $��3. Bundles.xlsx / xls ->��6   $
��Left=�63   $��	22,85,0,0q�6N   $	��Top=�6a   $	��131��6y   �� 5   
   -   5        $��Convert bundle��5      $
��Left=�6*   $��
291,88,0,0q�6E   $	��Top=�6Y   $��90��6�   �� 5   
   -   5        $ name_Sheet_Bundle��5      $
��Left=�6-   $��
174,89,0,0q�6H   $
 Wrap=�6\   $	��Top=�6p   $	��104��6�   $ fill which spreadsheet��6�   �� 5   
   -   5        $ image_ok_bundles��5      $
��Left=�6*   $��19��6E   $��
387,88,0,0q�6Q   $	��Top=�6e   $��26��6}   $��/img/image.jpg��6�   $
 Both=�6�   $y�Hidden=�6�   �� 5   
   $��4. Save txt file��6   $
��Left=�6+   $��
23,115,0,0q�6F   $	��Top=�6Z   $��87��6r   �� 5   
   -	   5        $��	Save file��5      $
��Left=�6%   $��175,119,0,0q�6@   $	��Top=�6U   6m   �� 5   
   -
   5        $ image_ok_save_file��5      $
��Left=�6,   $��19��6G   $��233,120,0,0q�6S   $	��Top=�6h   $��26��6�   $��/img/image.jpg��6�   $
 Both=�6�   $y�Hidden=�6�   �� 5   
   $��5. Exit from application��6   $
��Left=�63   $��
23,144,0,0q�6N   $	��Top=�6b   $	��135��6z   $��26��6�   �� 5   
   -   5        $
��Exit��5      $
��Left=�6    $��175,148,0,0q�6;   $	��Top=�6P   6h   5&      5*      