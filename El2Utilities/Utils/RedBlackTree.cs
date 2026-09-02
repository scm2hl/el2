using System;

namespace El2Core.Utils
{
    public sealed class RedBlackTree : Tree
    {
        private Color Black = Color.Black;
        private Color Red = Color.Red;
        private Node parentNode;
        private Node grandParentNode;
        private Node tempNode;

        /// <summary>
        /// Inserts an item into the Red-Black Tree while maintaining its properties.
        /// </summary>
        /// <param name="item"></param>
        public void Insert(IComparable item)
        {
            //root.data = item;
            currentNode = parentNode = grandParentNode = root;
            freshNode.data = item;
            int returnedValue = 0;
            while (Compare(item, currentNode) != 0)
            {
                tempNode = grandParentNode;
                grandParentNode = parentNode;
                parentNode = currentNode;
                returnedValue = Compare(item, currentNode);
                if (returnedValue < 0) currentNode = currentNode.left;
                else currentNode = currentNode.right;
                if (currentNode.left.color == Color.Red && currentNode.right.color == Color.Red) ReArrange(item);
            }
            if (currentNode == freshNode)
            {
                currentNode = new Node(item, freshNode, freshNode);
                if (Compare(item, parentNode) < 0) parentNode.left = currentNode;
                else parentNode.right = currentNode;
                ReArrange(item);
            }
        }
        /// <summary>
        /// Rearranges the tree to maintain Red-Black Tree properties after an insertion.
        /// </summary>
        /// <param name="item"></param>
        private void ReArrange(IComparable item)
        {
            currentNode.color = Red;
            currentNode.left.color = Color.Black;
            currentNode.right.color = Color.Black;
            if (parentNode.color == Color.Red)
            {
                grandParentNode.color = Color.Red;
                bool compareWithGrandParentNode = (Compare(item, grandParentNode) < 0);
                bool compareWithParentNode = (Compare(item, parentNode) < 0);
                if (compareWithGrandParentNode != compareWithParentNode) parentNode = Rotate(item, grandParentNode);
                currentNode = Rotate(item, tempNode);
                currentNode.color = Black;
            }
            root.right.color = Color.Black;
        }
        private Node Rotate(IComparable item, Node parentNode)
        {
            int value;
            if (Compare(item, parentNode) < 0)
            {
                value = Compare(item, parentNode.left);
                if (value < 0) parentNode.left = this.Rotate(parentNode.left, Direction.Left);
                else parentNode.left = this.Rotate(parentNode.left, Direction.Right);
                return parentNode.left;
            }
            else
            {
                value = Compare(item, parentNode.right);
                if (value < 0) parentNode.right = this.Rotate(parentNode.right, Direction.Left);
                else parentNode.right = this.Rotate(parentNode.right, Direction.Right);
                return parentNode.right;
            }
        }
        private Node Rotate(Node node, Direction direction)
        {
            Node tempNode;
            if (direction == Direction.Left)
            {
                tempNode = node.left;
                node.left = tempNode.right;
                tempNode.right = node;
                return tempNode;
            }
            else
            {
                tempNode = node.right;
                node.right = tempNode.left;
                tempNode.left = node;
                return tempNode;
            }
        }
    }
    public enum Color
    {
        Red = 0, Black = 1
    }
    public enum Direction
    {
        Left,
        Right
    }

    public class Node
    {
        public IComparable data;
        public Node left;
        public Node right;
        public Color color = Color.Black;
        public Node(IComparable data) : this(data, null, null) { }
        public Node(IComparable data, Node left, Node right)
        {
            this.data = data;
            this.left = left;
            this.right = right;
        }
    }
    public class Tree
    {
        protected Node root;
        protected Node freshNode;
        protected Node currentNode;
        protected Tree()
        {
            freshNode = new Node(null);
            freshNode.left = freshNode.right = freshNode;
            root = new Node(null);
        }
        protected int Compare(IComparable item, Node node)
        {
            if (node != root) return item.CompareTo(node.data);
            else return 1;
        }
        public IComparable Search(IComparable data)
        {
            freshNode.data = data;
            currentNode = root.right;
            while (true)
            {
                if (Compare(data, currentNode) < 0) currentNode = currentNode.left;
                else if (Compare(data, currentNode) > 0) currentNode = currentNode.right;
                else if (currentNode != freshNode) return currentNode.data;
                else return null;
            }
        }
        protected void Display()
        {
            this.Display(root.right);
        }
        protected void Display(Node temp)
        {
            if (temp != freshNode)
            {
                Display(temp.left);
                Console.WriteLine(temp.data);
                Display(temp.right);
            }
        }
    }
}
