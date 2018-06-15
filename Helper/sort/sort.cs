using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

/// <summary>
/// 排序类库
/// </summary>
namespace Helper.sort
{
    public static class sort
    {

        /// <summary>
        /// 当某轮比较没有发生移动时，就可以确定排序完成了
        /// <para>稳定排序</para>
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="arr"></param>
        public static void BubbleSortAscending1<T>(this IList<T> arr)
            where T : IComparable
        {
            bool exchanges;
            do
            {
                exchanges = false;
                for (int i = 0; i < arr.Count - 1; i++)
                {
                    if (arr[i].CompareTo(arr[i + 1]) > 0)
                    {
                        T temp = arr[i];
                        arr[i] = arr[i + 1];
                        arr[i + 1] = temp;
                        //当某轮比较没有发生移动时，就可以确定排序完成了
                        //否则应该继续排序
                        exchanges = true;
                    }
                }
            } while (exchanges);
        }

   
        /// <summary>
        /// 快速排序算法
        /// <para>不稳定排序</para>
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="arr"></param>
        private static void QuickSortAscending1<T>(this IList<T> arr)
            where T : IComparable
        {
            QuickSortAscending1_Do(arr, 0, 0, arr.Count - 1);
        }
        private static void QuickSortAscending1_Do<T>(IList<T> arr, int indexOfRightPlaceToFind, int first, int last)
            where T : IComparable
        {
            if (first < last)
            {
                int rightPlace = QuickSortAscending1_Find(indexOfRightPlaceToFind, arr, first, last);
                if (first + 1 < last)
                {
                    QuickSortAscending1_Do(arr, first, first, rightPlace - 1);
                    QuickSortAscending1_Do(arr, rightPlace + 1, rightPlace + 1, last);
                }
            }
        }
        private static int QuickSortAscending1_Find<T>(int indexOfRightPlaceToFind, IList<T> arr, int first, int last)
            where T : IComparable
        {
            bool searchRight = true;
            int indexOfLeftSearch = first;
            int indexOfRightSearch = last;
            do
            {
                if (searchRight)
                {
                    while (arr[indexOfRightPlaceToFind].CompareTo(arr[indexOfRightSearch]) <= 0)
                    {
                        indexOfRightSearch--;
                        if (indexOfRightPlaceToFind == indexOfRightSearch)
                        {
                            searchRight = false;
                            break;
                        }
                    }
                    if (searchRight)
                    {
                        T temp = arr[indexOfRightPlaceToFind];
                        arr[indexOfRightPlaceToFind] = arr[indexOfRightSearch];
                        arr[indexOfRightSearch] = temp;
                        indexOfRightPlaceToFind = indexOfRightSearch;
                        searchRight = false;
                    }
                }
                else
                {
                    while (arr[indexOfRightPlaceToFind].CompareTo(arr[indexOfLeftSearch]) >= 0)
                    {
                        indexOfLeftSearch++;
                        if (indexOfRightPlaceToFind == indexOfLeftSearch)
                        {
                            searchRight = true;
                            break;
                        }
                    }
                    if (!searchRight)
                    {
                        T temp = arr[indexOfRightPlaceToFind];
                        arr[indexOfRightPlaceToFind] = arr[indexOfLeftSearch];
                        arr[indexOfLeftSearch] = temp;
                        indexOfRightPlaceToFind = indexOfLeftSearch;
                        searchRight = true;
                    }
                }
            } while (indexOfLeftSearch < indexOfRightPlaceToFind || indexOfRightPlaceToFind < indexOfRightSearch);
            return indexOfRightPlaceToFind;
        }
    }
}
