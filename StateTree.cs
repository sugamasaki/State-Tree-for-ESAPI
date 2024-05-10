using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Collections.Generic;
using System.Reflection;
using System.Runtime.CompilerServices;
using VMS.TPS.Common.Model.API;
using VMS.TPS.Common.Model.Types;
using System.Collections;
using System.Xml.Linq;

[assembly: AssemblyVersion("1.0.0.1")]

namespace VMS.TPS
{
    public class Script
    {
        public Script()
        {
        }

        [MethodImpl(MethodImplOptions.NoInlining)]
        public void Execute(ScriptContext context) //, System.Windows.Window window, ScriptEnvironment environment)
        {
            // TODO : Add here the code that is called when the script is launched from Eclipse.
            StateTree stateTree = new StateTree(context);
            stateTree.ShowDialog();
        }
    }

    public partial class StateTree : Window
    {
        private List<Type> valueType = new List<Type>() { typeof(bool), typeof(sbyte), typeof(byte), typeof(short), typeof(ushort), typeof(int), typeof(uint), typeof(long), typeof(ulong), typeof(char), typeof(float), typeof(double), typeof(decimal), typeof(string) };
        XElement xmlAPI;
        private TextBox textBox;
        Object GetPropertyValue(Object obj, string name)
        {
            var property = obj.GetType().GetProperty(name);
            return property.GetValue(obj);
        }
        public StateTree(ScriptContext context)
        {
            string versionInfo = "11";
            string filePath = "";
            if (context.GetType().GetProperties().Any(p => p.Name == "VersionInfo"))
            {
                versionInfo = (string)GetPropertyValue(context, "VersionInfo");
            }
            if (versionInfo == "11")
            {
                filePath = @"C:\Program Files (x86)\Varian\Vision\11.0\Bin64\VMS.TPS.Common.Model.API.xml";
            }
            else
            {
                string rtmPath = @"C:\Program Files (x86)\Varian\RTM\";
                if (!Directory.Exists(rtmPath))
                {
                    rtmPath = @"C:\Program Files\Varian\RTM\";
                }
                string[] versionPaths = Directory.GetDirectories(rtmPath);
                foreach (var versionPath in versionPaths)
                {
                    if (versionInfo.Contains(versionPath.Split('\\').Last()))
                    {
                        filePath = versionPath + @"\esapi\API\VMS.TPS.Common.Model.API.xml";
                    }
                }
            }
            if (File.Exists(filePath))
            {
                xmlAPI = XElement.Load(filePath);
            }
            else
            {
                MessageBox.Show(filePath + "is not found.");
            }
            TreeView treeView = CreateTreeView(context);
            InitializeComponent(treeView);
        }
        private void InitializeComponent(TreeView treeView)
        {
            this.Title = "StateTree for ESAPI";
            Grid grid = new Grid();
            grid.ColumnDefinitions.Add(new ColumnDefinition());
            grid.ColumnDefinitions.Add(new ColumnDefinition());
            grid.ColumnDefinitions.Add(new ColumnDefinition());
            grid.ColumnDefinitions[0].Width = new GridLength(2.0, GridUnitType.Star);
            grid.ColumnDefinitions[1].Width = GridLength.Auto;
            grid.ColumnDefinitions[2].Width = new GridLength(1.0, GridUnitType.Star);
            treeView.SetValue(Grid.ColumnProperty, 0);
            GridSplitter splitter = new GridSplitter();
            splitter.SetValue(Grid.ColumnProperty, 1);
            splitter.Width = 5;
            splitter.HorizontalAlignment = HorizontalAlignment.Center;
            textBox = new TextBox();
            textBox.VerticalScrollBarVisibility = ScrollBarVisibility.Visible;
            textBox.TextWrapping = TextWrapping.Wrap;
            textBox.SetValue(Grid.ColumnProperty, 2);
            grid.Children.Add(treeView);
            grid.Children.Add(splitter);
            grid.Children.Add(textBox);
            this.Content = grid;
        }
        private TreeView CreateTreeView(Object obj)
        {
            TreeView treeView = new TreeView();
            TreeViewItem rootItem = CreateTreeViewItem(obj);
            treeView.Items.Add(rootItem);
            return treeView;
        }

        private TreeViewItem CreateTreeViewItem(Object obj)
        {
            TreeViewItem item = new TreeViewItem();
            item.Header = obj.GetType().Name;
            item.Tag = obj;
            item.Items.Add("Loading..."); // dummy
            item.Expanded += OnTreeViewItemExpanded;
            item.Selected += OnItemSelected;
            return item;
        }

        private void OnTreeViewItemExpanded(object sender, RoutedEventArgs e)
        {
            TreeViewItem item = sender as TreeViewItem;
            if (item.Items.Count == 1 && item.Items[0].ToString() == "Loading...")
            {
                item.Items.Clear(); // delete dummy placeholder
                Object obj = item.Tag;
                CreateItems(obj, item);
            }
        }

        private void OnFieldItemExpanded(object sender, RoutedEventArgs e)
        {
            TreeViewItem item = sender as TreeViewItem;
            if (item.Items.Count == 1 && item.Items[0] is FieldInfo)
            {
                Object obj = item.Tag;
                FieldInfo field = item.Items[0] as FieldInfo;
                item.Items.Clear();
                item.Tag = field;
                try
                {
                    var value = field.GetValue(obj);
                    if (value.GetType() != obj.GetType())
                    {
                        if (field.IsInitOnly || field.IsLiteral)
                        {
                            CreateFildOrPropertyItem(value, field, item, true, false);
                        }
                        else
                        {
                            CreateFildOrPropertyItem(value, field, item);
                        }
                    }

                }
                catch (NullReferenceException)
                {
                    item.Header = item.Header.ToString() + ": null";
                }
                catch (TargetInvocationException)
                {
                    item.Header = item.Header.ToString() + ": null";
                }
                catch (TargetParameterCountException)
                {
                    item.Header = item.Header.ToString() + " TargetParameterCountException Error?";
                }
                catch (Exception ex)
                {
                    MessageBox.Show(string.Format("parent:{0} {1} {2}\n{3}", item.Header, field.Name, field.GetType().Name, ex.ToString()));
                }
            }
        }

        private void OnPropertyItemExpanded(object sender, RoutedEventArgs e)
        {
            TreeViewItem item = sender as TreeViewItem;
            if (item.Items.Count == 1 && item.Items[0] is PropertyInfo)
            {
                Object obj = item.Tag;
                PropertyInfo property = item.Items[0] as PropertyInfo;
                item.Items.Clear();
                item.Tag = property;
                try
                {
                    if (property.GetIndexParameters().Length == 0)
                    {
                        var value = property.GetValue(obj);
                        CreateFildOrPropertyItem(value, property, item, property.CanRead, property.CanWrite);
                    }
                    else if (property.GetIndexParameters().Length == 1)
                    {
                        var parameters = property.GetIndexParameters();
                        if (parameters[0].ParameterType == typeof(int))
                        {
                            var count = (int)GetPropertyValue(obj, "Count");
                            for (int i = 0; i < count; i++)
                            {
                                var item1 = new TreeViewItem();
                                item1.Header = property.Name + i.ToString();
                                var value = property.GetValue(obj, new object[] { i });
                                item1.Tag = value;
                                item1.Selected += OnItemSelected;
                                item.Items.Add(item1);
                                CreateFildOrPropertyItem(value, property, item1, property.CanRead, property.CanWrite);
                            }
                        }
                        else if (parameters[0].ParameterType == typeof(string))
                        {
                            var keys = (IEnumerable<string>)GetPropertyValue(obj, "Keys");
                            int count = 0;
                            foreach (var key in keys)
                            {
                                var item1 = new TreeViewItem();
                                item1.Header = property.Name + count.ToString();
                                var value = property.GetValue(obj, new object[] { key });
                                item1.Tag = value;
                                item1.Selected += OnItemSelected;
                                item.Items.Add(item1);
                                CreateFildOrPropertyItem(value, property, item1, property.CanRead, property.CanWrite);
                                count++;
                            }
                        }
                        else
                        {
                            item.Header = item.Header.ToString() + " Error! this type is not supported yet.";
                        }
                    }
                    else
                    {
                        item.Header = item.Header.ToString() + " Error! this type is not supported yet.";
                    }
                }
                catch (NullReferenceException)
                {
                    item.Header = item.Header.ToString() + ": null";
                }
                catch (TargetInvocationException)
                {
                    item.Header = item.Header.ToString() + ": null";
                }
                catch (TargetParameterCountException)
                {
                    item.Header = item.Header.ToString() + " Error! TargetParameterCountException";
                }
                catch (Exception ex)
                {
                    MessageBox.Show(string.Format("parent:{0} {1} {2}\n{3}", item.Header, property.Name, property.GetType().Name, ex.ToString()));
                }
            }

        }
        private void OnItemSelected(object sender, RoutedEventArgs e)
        {
            TreeViewItem item = sender as TreeViewItem;
            textBox.Text = "";
            if (item.Tag == null)
            {
                e.Handled = true;
                return;
            }
            if (item.Items.Count == 1 && (item.Items[0] is PropertyInfo || item.Items[0] is FieldInfo))
            {
                string name = "";
                if (item.Items[0] is PropertyInfo)
                {
                    name = ((PropertyInfo)item.Items[0]).PropertyType.FullName;
                }
                else if (item.Items[0] is FieldInfo)
                {
                    name = ((FieldInfo)item.Items[0]).FieldType.FullName;

                }
                if (xmlAPI == null)
                {
                    e.Handled = true;
                    return;
                }
                IEnumerable<XElement> members = xmlAPI.Element("members").Elements("member");
                if (members.Any(m => m.Attribute("name").Value.EndsWith(name)))
                {
                    XElement member = members.Single(m => m.Attribute("name").Value.EndsWith(name));
                    foreach (var attribute in member.Attributes())
                    {
                        textBox.Text += attribute.Value;
                    }
                    foreach (var element in member.Elements())
                    {
                        textBox.Text += element.Value;
                    }
                }
                else
                {
                    TreeViewItem parentItem = item.Parent as TreeViewItem;
                    if (parentItem.Tag is PropertyInfo)
                    {
                        name = ((PropertyInfo)parentItem.Tag).PropertyType.FullName + "." + item.Header;
                    }
                    else if (parentItem.Tag is FieldInfo)
                    {
                        name = ((FieldInfo)parentItem.Tag).FieldType.FullName + "." + item.Header;
                    }
                    else
                    {
                        name = parentItem.Tag.GetType().FullName + "." + item.Header;
                    }
                    if (members.Any(m => m.Attribute("name").Value.EndsWith(name)))
                    {
                        XElement member = members.Single(m => m.Attribute("name").Value.EndsWith(name));
                        foreach (var attribute in member.Attributes())
                        {
                            textBox.Text += attribute.Value;
                        }
                        foreach (var element in member.Elements())
                        {
                            textBox.Text += element.Value;
                        }
                    }
                }
            }
            else
            {
                string name = "";
                if (item.Tag is PropertyInfo)
                {
                    name = ((PropertyInfo)item.Tag).PropertyType.FullName;
                }
                else if (item.Tag is FieldInfo)
                {
                    name = ((FieldInfo)item.Tag).FieldType.FullName;

                }
                else if (item.Tag is MethodInfo)
                {
                    MethodInfo method = item.Tag as MethodInfo;
                    if (method.GetParameters().Length == 0)
                    {
                        name += string.Format(".{0}()", method.Name);
                    }
                    else
                    {
                        foreach (var p in method.GetParameters())
                        {
                            if (name == "")
                            {
                                name += string.Format(".{0}({1}", method.Name, p.ParameterType.FullName);
                            }
                            else
                            {
                                name += string.Format(",{0}", p.ParameterType.FullName);
                            }
                        }
                        name += ")";
                    }
                    TreeViewItem parentItem = item.Parent as TreeViewItem;
                    if (parentItem.Tag is PropertyInfo)
                    {
                        name = ((PropertyInfo)parentItem.Tag).PropertyType.FullName + name;
                    }
                    else if (parentItem.Tag is FieldInfo)
                    {
                        name = ((FieldInfo)parentItem.Tag).FieldType.FullName + name;
                    }
                }
                else
                {
                    name = item.Tag.GetType().FullName;
                }
                if (xmlAPI == null)
                {
                    e.Handled = true;
                    return;
                }
                IEnumerable<XElement> members = xmlAPI.Element("members").Elements("member");
                if (members.Any(m => m.Attribute("name").Value.EndsWith(name)))
                {
                    XElement member = members.Single(m => m.Attribute("name").Value.EndsWith(name));
                    foreach (var attribute in member.Attributes())
                    {
                        textBox.Text += attribute.Value;
                    }
                    foreach (var element in member.Elements())
                    {
                        textBox.Text += element.Value;
                    }
                }
                else
                {
                    TreeViewItem parentItem = item.Parent as TreeViewItem;
                    if (parentItem.Tag is PropertyInfo)
                    {
                        name = ((PropertyInfo)parentItem.Tag).PropertyType.FullName + "." + item.Header;
                    }
                    else if (parentItem.Tag is FieldInfo)
                    {
                        name = ((FieldInfo)parentItem.Tag).FieldType.FullName + "." + item.Header;
                    }
                    else
                    {
                        name = parentItem.Tag.GetType().FullName + "." + item.Header;
                    }
                    if (members.Any(m => m.Attribute("name").Value.EndsWith(name)))
                    {
                        XElement member = members.Single(m => m.Attribute("name").Value.EndsWith(name));
                        foreach (var attribute in member.Attributes())
                        {
                            textBox.Text += attribute.Value;
                        }
                        foreach (var element in member.Elements())
                        {
                            textBox.Text += element.Value;
                        }
                    }
                }
            }
            e.Handled = true;
        }
        private void CreateItems(Object obj, TreeViewItem parentItem)
        {
            if (obj == null)
            {
                parentItem.Header = parentItem.Header.ToString() + ": null";
                return;
            }
            var type = obj.GetType();
            CreateFieldItems(obj, type, parentItem);
            CreatePropertyItems(obj, type, parentItem);
            CreateMethodItems(obj, type, parentItem);
        }
        private void CreateFieldItems(Object obj, Type objType, TreeViewItem parentItem)
        {
            foreach (var field in objType.GetFields())
            {
                var item = new TreeViewItem();
                item.Header = field.Name;
                item.Tag = obj;
                item.Items.Add(field);
                item.Expanded += OnFieldItemExpanded;
                item.Selected += OnItemSelected;
                parentItem.Items.Add(item);

            }
        }

        private void CreatePropertyItems(Object obj, Type objType, TreeViewItem parentItem)
        {
            foreach (var property in objType.GetProperties())
            {
                var item = new TreeViewItem();
                item.Header = property.Name;
                item.Tag = obj;
                item.Items.Add(property);
                item.Expanded += OnPropertyItemExpanded;
                item.Selected += OnItemSelected;
                parentItem.Items.Add(item);
            }
        }
        private void CreateFildOrPropertyItem(Object value, Object info, TreeViewItem parentItem, bool canGet = true, bool canSet = true)
        {
            string name = "";
            if (info is PropertyInfo)
            {
                name = ((PropertyInfo)info).Name;
            }
            else if (info is FieldInfo)
            {
                name = ((FieldInfo)info).Name;
            }
            var type = value.GetType();
            if (type.IsArray)
            {
                ArrayFor(value as Array, name, parentItem, 0, canGet, canSet);
            }
            else if (type.FullName.StartsWith("System.Collections") || type.Name.Contains("d__"))
            {
                var enumerable = value as IEnumerable;
                int count = 0;
                foreach (var v in enumerable)
                {
                    var item = new TreeViewItem();
                    item.Header = string.Format("{0}[{1}]", name, count);
                    item.Tag = v;
                    item.Selected += OnItemSelected;
                    parentItem.Items.Add(item);
                    count++;
                    if (valueType.Contains(v.GetType()))
                    {
                        CreateChildItems(v, item, canGet, canSet);
                    }
                    else
                    {
                        CreateItems(v, item);
                    }
                }
            }
            else if (valueType.Contains(type) || type.IsEnum)
            {
                CreateChildItems(value, parentItem, canGet, canSet);
            }
            else
            {
                CreateItems(value, parentItem);
            }
        }
        private void CreateMethodItems(Object obj, Type objType, TreeViewItem parentItem)
        {
            List<string> property_methods = new List<string>();
            foreach (var property in objType.GetProperties())
            {
                if (property.CanRead)
                {
                    property_methods.Add("get_" + property.Name);
                }
                if (property.CanWrite)
                {
                    property_methods.Add("set_" + property.Name);
                }
            }
            foreach (var method in objType.GetMethods())
            {
                if (property_methods.Contains(method.Name))
                {
                    continue;
                }
                var item = new TreeViewItem();
                item.Tag = method;
                item.Selected += OnItemSelected;
                item.Header = "";
                if (method.GetParameters().Length == 0)
                {
                    item.Header = string.Format("{0} {1}()", method.ReturnType.Name, method.Name);
                }
                else
                {
                    foreach (var p in method.GetParameters())
                    {
                        if (item.Header.ToString() == "")
                        {
                            item.Header = string.Format("{0} {1}({2} {3}", method.ReturnType.Name, method.Name, p.ParameterType.Name, p.Name);
                        }
                        else
                        {
                            item.Header = string.Format("{0}, {1} {2}", item.Header, p.ParameterType.Name, p.Name);
                        }
                    }
                    item.Header = item.Header.ToString() + ")";
                }
                parentItem.Items.Add(item);
            }
        }

        private void ArrayFor(Array arr, string name, TreeViewItem parentItem, int dimension, bool canGet, bool canSet)
        {
            var length = arr.GetLength(dimension);
            if (dimension == arr.Rank - 1)
            {
                for (int i = 0; i < length; i++)
                {
                    var item = new TreeViewItem();
                    int[] indices = new int[arr.Rank];
                    var strs = name.Split(',');
                    item.Header = strs[0] + "[";
                    item.Selected += OnItemSelected;
                    for (int j = 0; j < arr.Rank - 1; j++)
                    {
                        item.Header = string.Format("{0},{1}", item.Header, strs[j + 1]);
                        indices[j] = Int32.Parse(strs[j + 1]);
                    }
                    indices[arr.Rank - 1] = i;
                    if (arr.Rank == 1)
                    {
                        item.Header = string.Format("{0}{1}]", item.Header, i);
                    }
                    else
                    {
                        item.Header = string.Format("{0},{1}]", item.Header, i);
                    }
                    parentItem.Items.Add(item);
                    if (valueType.Contains(arr.GetType().GetElementType()))
                    {
                        item.Tag = arr.GetValue(indices);
                        CreateChildItems(arr.GetValue(indices), item, canGet, canSet);
                    }
                    else if (arr.GetType().GetElementType().IsArray)
                    {
                        ArrayFor(arr.GetValue(indices) as Array, strs[0], item, 0, canGet, canSet);
                    }
                    else
                    {
                        item.Tag = arr.GetValue(indices);
                        CreateItems(arr.GetValue(indices), item);
                    }
                }
            }
            else
            {
                for (int i = 0; i < length; i++)
                {
                    ArrayFor(arr, string.Format("{0},{1}", name, i), parentItem, dimension + 1, canGet, canSet);
                }
            }
        }

        private void CreateChildItems(Object obj, TreeViewItem parentItem, bool canGet = true, bool canSet = true)
        {
            var item1 = new TreeViewItem();
            var item2 = new TreeViewItem();
            string type = "type: " + obj.GetType().Name + " {";
            if (canGet)
            {
                type += "get;";
            }
            if (canSet)
            {
                type += "set;";
            }
            type += "}";
            item1.Header = type;
            item2.Header = "value: " + obj.ToString();
            item1.Selected += OnItemSelected;
            item2.Selected += OnItemSelected;
            parentItem.Items.Add(item1);
            parentItem.Items.Add(item2);
        }
    }
}
