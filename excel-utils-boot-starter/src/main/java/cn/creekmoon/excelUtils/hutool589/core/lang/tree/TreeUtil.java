package cn.creekmoon.excelUtils.hutool589.core.lang.tree;

import cn.creekmoon.excelUtils.hutool589.core.collection.IterUtil;
import cn.creekmoon.excelUtils.hutool589.core.lang.tree.parser.DefaultNodeParser;
import cn.creekmoon.excelUtils.hutool589.core.lang.tree.parser.NodeParser;
import cn.creekmoon.excelUtils.hutool589.core.util.ObjectUtil;

import java.util.ArrayList;
import java.util.List;
import java.util.Map;

/**
 * 树工具类
 *
 * @author liangbaikai
 */
public class TreeUtil {

	/**
	 * 构建单root节点树
	 *
	 * @param list 源数据集合
	 * @return {@link cn.creekmoon.excelUtils.hutool589.core.lang.tree.Tree}
	 * @since 5.7.2
	 */
	public static cn.creekmoon.excelUtils.hutool589.core.lang.tree.Tree<Integer> buildSingle(List<cn.creekmoon.excelUtils.hutool589.core.lang.tree.TreeNode<Integer>> list) {
		return buildSingle(list, 0);
	}

	/**
	 * 树构建
	 *
	 * @param list 源数据集合
	 * @return List
	 */
	public static List<cn.creekmoon.excelUtils.hutool589.core.lang.tree.Tree<Integer>> build(List<cn.creekmoon.excelUtils.hutool589.core.lang.tree.TreeNode<Integer>> list) {
		return build(list, 0);
	}

	/**
	 * 构建单root节点树<br>
	 * 它会生成一个以指定ID为ID的空的节点，然后逐级增加子节点。
	 *
	 * @param <E>      ID类型
	 * @param list     源数据集合
	 * @param parentId 最顶层父id值 一般为 0 之类
	 * @return {@link cn.creekmoon.excelUtils.hutool589.core.lang.tree.Tree}
	 * @since 5.7.2
	 */
	public static <E> cn.creekmoon.excelUtils.hutool589.core.lang.tree.Tree<E> buildSingle(List<cn.creekmoon.excelUtils.hutool589.core.lang.tree.TreeNode<E>> list, E parentId) {
		return buildSingle(list, parentId, TreeNodeConfig.DEFAULT_CONFIG, new DefaultNodeParser<>());
	}

	/**
	 * 树构建
	 *
	 * @param <E>      ID类型
	 * @param list     源数据集合
	 * @param parentId 最顶层父id值 一般为 0 之类
	 * @return List
	 */
	public static <E> List<cn.creekmoon.excelUtils.hutool589.core.lang.tree.Tree<E>> build(List<TreeNode<E>> list, E parentId) {
		return build(list, parentId, TreeNodeConfig.DEFAULT_CONFIG, new DefaultNodeParser<>());
	}

	/**
	 * 构建单root节点树<br>
	 * 它会生成一个以指定ID为ID的空的节点，然后逐级增加子节点。
	 *
	 * @param <T>        转换的实体 为数据源里的对象类型
	 * @param <E>        ID类型
	 * @param list       源数据集合
	 * @param parentId   最顶层父id值 一般为 0 之类
	 * @param nodeParser 转换器
	 * @return {@link cn.creekmoon.excelUtils.hutool589.core.lang.tree.Tree}
	 * @since 5.7.2
	 */
	public static <T, E> cn.creekmoon.excelUtils.hutool589.core.lang.tree.Tree<E> buildSingle(List<T> list, E parentId, NodeParser<T, E> nodeParser) {
		return buildSingle(list, parentId, TreeNodeConfig.DEFAULT_CONFIG, nodeParser);
	}

	/**
	 * 树构建
	 *
	 * @param <T>        转换的实体 为数据源里的对象类型
	 * @param <E>        ID类型
	 * @param list       源数据集合
	 * @param parentId   最顶层父id值 一般为 0 之类
	 * @param nodeParser 转换器
	 * @return List
	 */
	public static <T, E> List<cn.creekmoon.excelUtils.hutool589.core.lang.tree.Tree<E>> build(List<T> list, E parentId, NodeParser<T, E> nodeParser) {
		return build(list, parentId, TreeNodeConfig.DEFAULT_CONFIG, nodeParser);
	}

	/**
	 * 树构建
	 *
	 * @param <T>            转换的实体 为数据源里的对象类型
	 * @param <E>            ID类型
	 * @param list           源数据集合
	 * @param rootId         最顶层父id值 一般为 0 之类
	 * @param treeNodeConfig 配置
	 * @param nodeParser     转换器
	 * @return List
	 */
	public static <T, E> List<cn.creekmoon.excelUtils.hutool589.core.lang.tree.Tree<E>> build(List<T> list, E rootId, TreeNodeConfig treeNodeConfig, NodeParser<T, E> nodeParser) {
		return buildSingle(list, rootId, treeNodeConfig, nodeParser).getChildren();
	}

	/**
	 * 构建单root节点树<br>
	 * 它会生成一个以指定ID为ID的空的节点，然后逐级增加子节点。
	 *
	 * @param <T>            转换的实体 为数据源里的对象类型
	 * @param <E>            ID类型
	 * @param list           源数据集合
	 * @param rootId         最顶层父id值 一般为 0 之类
	 * @param treeNodeConfig 配置
	 * @param nodeParser     转换器
	 * @return {@link cn.creekmoon.excelUtils.hutool589.core.lang.tree.Tree}
	 * @since 5.7.2
	 */
	public static <T, E> cn.creekmoon.excelUtils.hutool589.core.lang.tree.Tree<E> buildSingle(List<T> list, E rootId, TreeNodeConfig treeNodeConfig, NodeParser<T, E> nodeParser) {
		return TreeBuilder.of(rootId, treeNodeConfig)
				.append(list, rootId, nodeParser).build();
	}

	/**
	 * 树构建，按照权重排序
	 *
	 * @param <E>    ID类型
	 * @param map    源数据Map
	 * @param rootId 最顶层父id值 一般为 0 之类
	 * @return List
	 * @since 5.6.7
	 */
	public static <E> List<cn.creekmoon.excelUtils.hutool589.core.lang.tree.Tree<E>> build(Map<E, cn.creekmoon.excelUtils.hutool589.core.lang.tree.Tree<E>> map, E rootId) {
		return buildSingle(map, rootId).getChildren();
	}

	/**
	 * 单点树构建，按照权重排序<br>
	 * 它会生成一个以指定ID为ID的空的节点，然后逐级增加子节点。
	 *
	 * @param <E>    ID类型
	 * @param map    源数据Map
	 * @param rootId 根节点id值 一般为 0 之类
	 * @return {@link cn.creekmoon.excelUtils.hutool589.core.lang.tree.Tree}
	 * @since 5.7.2
	 */
	public static <E> cn.creekmoon.excelUtils.hutool589.core.lang.tree.Tree<E> buildSingle(Map<E, cn.creekmoon.excelUtils.hutool589.core.lang.tree.Tree<E>> map, E rootId) {
		final cn.creekmoon.excelUtils.hutool589.core.lang.tree.Tree<E> tree = IterUtil.getFirstNoneNull(map.values());
		if (null != tree) {
			final TreeNodeConfig config = tree.getConfig();
			return TreeBuilder.of(rootId, config)
					.append(map)
					.build();
		}

		return createEmptyNode(rootId);
	}

	/**
	 * 获取ID对应的节点，如果有多个ID相同的节点，只返回第一个。<br>
	 * 此方法只查找此节点及子节点，采用递归深度优先遍历。
	 *
	 * @param <T>  ID类型
	 * @param node 节点
	 * @param id   ID
	 * @return 节点
	 * @since 5.2.4
	 */
	public static <T> cn.creekmoon.excelUtils.hutool589.core.lang.tree.Tree<T> getNode(cn.creekmoon.excelUtils.hutool589.core.lang.tree.Tree<T> node, T id) {
		if (ObjectUtil.equal(id, node.getId())) {
			return node;
		}

		final List<cn.creekmoon.excelUtils.hutool589.core.lang.tree.Tree<T>> children = node.getChildren();
		if (null == children) {
			return null;
		}

		// 查找子节点
		cn.creekmoon.excelUtils.hutool589.core.lang.tree.Tree<T> childNode;
		for (cn.creekmoon.excelUtils.hutool589.core.lang.tree.Tree<T> child : children) {
			childNode = child.getNode(id);
			if (null != childNode) {
				return childNode;
			}
		}

		// 未找到节点
		return null;
	}

	/**
	 * 获取所有父节点名称列表
	 *
	 * <p>
	 * 比如有个人在研发1部，他上面有研发部，接着上面有技术中心<br>
	 * 返回结果就是：[研发一部, 研发中心, 技术中心]
	 *
	 * @param <T>                节点ID类型
	 * @param node               节点
	 * @param includeCurrentNode 是否包含当前节点的名称
	 * @return 所有父节点名称列表，node为null返回空List
	 * @since 5.2.4
	 */
	public static <T> List<CharSequence> getParentsName(cn.creekmoon.excelUtils.hutool589.core.lang.tree.Tree<T> node, boolean includeCurrentNode) {
		final List<CharSequence> result = new ArrayList<>();
		if (null == node) {
			return result;
		}

		if (includeCurrentNode) {
			result.add(node.getName());
		}

		cn.creekmoon.excelUtils.hutool589.core.lang.tree.Tree<T> parent = node.getParent();
		while (null != parent) {
			result.add(parent.getName());
			parent = parent.getParent();
		}
		return result;
	}

	/**
	 * 创建空Tree的节点
	 *
	 * @param id  节点ID
	 * @param <E> 节点ID类型
	 * @return {@link cn.creekmoon.excelUtils.hutool589.core.lang.tree.Tree}
	 * @since 5.7.2
	 */
	public static <E> cn.creekmoon.excelUtils.hutool589.core.lang.tree.Tree<E> createEmptyNode(E id) {
		return new Tree<E>().setId(id);
	}
}
