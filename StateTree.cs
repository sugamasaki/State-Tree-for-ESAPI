using System;
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
            //StateTree stateTree = new StateTree(context.PlanSetup);
            //StateTree stateTree = new StateTree(context.StructureCodes);
            stateTree.ShowDialog();
        }
    }

    public partial class StateTree : Window
    {
        private List<Type> valueType = new List<Type>() { typeof(bool), typeof(sbyte), typeof(byte), typeof(short), typeof(ushort), typeof(int), typeof(uint), typeof(long), typeof(ulong), typeof(char), typeof(float), typeof(double), typeof(decimal), typeof(string) };
        Object GetPropertyValue(Object obj, string name)
        {
            var property = obj.GetType().GetProperty(name);
            return property.GetValue(obj);
        }
        public StateTree(Object obj)
        {
            TreeView treeView = CreateTreeView(obj);
            InitializeComponent(treeView);
        }
        private void InitializeComponent(TreeView treeView)
        {
            this.Title = "StateTree for ESAPI";
            this.Content = treeView;
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
                try
                {
                    var value = field.GetValue(obj);
                    if (value.GetType() != obj.GetType())
                    {
                        CreateFildOrPropertyItem(value, field.Name, item);
                    }

                }
                catch (NullReferenceException)
                {
                    item.Header = item.Header.ToString() + ": null";
                }
                catch (TargetInvocationException)
                {
                }
                catch (TargetParameterCountException)
                {
                    item.Header = item.Header.ToString() + " TargetParameterCountException Error?";
                    //parentItem.Items.Remove(item);
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
                try
                {
                    if (property.GetIndexParameters().Length == 0)
                    {
                        var value = property.GetValue(obj);
                        CreateFildOrPropertyItem(value, property.Name, item, property.CanRead, property.CanWrite);
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
                                item.Items.Add(item1);
                                var value = property.GetValue(obj, new object[] { i });
                                CreateFildOrPropertyItem(value, property.Name, item1, property.CanRead, property.CanWrite);
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
                                item.Items.Add(item1);
                                var value = property.GetValue(obj, new object[] { key });
                                CreateFildOrPropertyItem(value, property.Name, item1, property.CanRead, property.CanWrite);
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
                }
                catch (TargetParameterCountException)
                {
                    item.Header = item.Header.ToString() + " TargetParameterCountException Error?";
                    //parentItem.Items.Remove(item);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(string.Format("parent:{0} {1} {2}\n{3}", item.Header, property.Name, property.GetType().Name, ex.ToString()));
                }
            }

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
                parentItem.Items.Add(item);
            }
        }
        private void CreateFildOrPropertyItem(Object value, string name, TreeViewItem parentItem, bool canGet = true, bool canSet = true)
        {
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
            else if (valueType.Contains(type))
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
                        item.Header = item.Header.ToString() + ")";
                    }
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
                        CreateChildItems(arr.GetValue(indices), item, canGet, canSet);
                    }
                    else if (arr.GetType().GetElementType().IsArray)
                    {
                        ArrayFor(arr.GetValue(indices) as Array, strs[0], item, 0, canGet, canSet);
                    }
                    else
                    {
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
            parentItem.Items.Add(item1);
            parentItem.Items.Add(item2);
        }
    }
}