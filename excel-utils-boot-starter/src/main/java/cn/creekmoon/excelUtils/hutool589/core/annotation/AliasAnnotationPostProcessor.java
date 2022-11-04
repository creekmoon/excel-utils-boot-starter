package cn.creekmoon.excelUtils.hutool589.core.annotation;

import cn.creekmoon.excelUtils.hutool589.core.lang.Assert;
import cn.creekmoon.excelUtils.hutool589.core.lang.Opt;
import cn.creekmoon.excelUtils.hutool589.core.map.ForestMap;
import cn.creekmoon.excelUtils.hutool589.core.map.LinkedForestMap;
import cn.creekmoon.excelUtils.hutool589.core.map.TreeEntry;
import cn.creekmoon.excelUtils.hutool589.core.util.ClassUtil;
import cn.creekmoon.excelUtils.hutool589.core.util.ObjectUtil;

import java.util.Map;

/**
 * <p>用于处理注解对象中带有{@link cn.creekmoon.excelUtils.hutool589.core.annotation.Alias}注解的属性。<br>
 * 当该处理器执行完毕后，{@link cn.creekmoon.excelUtils.hutool589.core.annotation.Alias}注解指向的目标注解的属性将会被包装并替换为
 * {@link ForceAliasedAnnotationAttribute}。
 *
 * @author huangchengxing
 * @see cn.creekmoon.excelUtils.hutool589.core.annotation.Alias
 * @see ForceAliasedAnnotationAttribute
 */
public class AliasAnnotationPostProcessor implements SynthesizedAnnotationPostProcessor {

	@Override
	public int order() {
		return Integer.MIN_VALUE;
	}

	@Override
	public void process(SynthesizedAnnotation synthesizedAnnotation, AnnotationSynthesizer synthesizer) {
		final Map<String, AnnotationAttribute> attributeMap = synthesizedAnnotation.getAttributes();

		// 记录别名与属性的关系
		final ForestMap<String, AnnotationAttribute> attributeAliasMappings = new LinkedForestMap<>(false);
		attributeMap.forEach((attributeName, attribute) -> {
			final String alias = Opt.ofNullable(attribute.getAnnotation(cn.creekmoon.excelUtils.hutool589.core.annotation.Alias.class))
					.map(Alias::value)
					.orElse(null);
			if (ObjectUtil.isNull(alias)) {
				return;
			}
			final AnnotationAttribute aliasAttribute = attributeMap.get(alias);
			Assert.notNull(aliasAttribute, "no method for alias: [{}]", alias);
			attributeAliasMappings.putLinkedNodes(alias, aliasAttribute, attributeName, attribute);
		});

		// 处理别名
		attributeMap.forEach((attributeName, attribute) -> {
			final AnnotationAttribute resolvedAttribute = Opt.ofNullable(attributeName)
					.map(attributeAliasMappings::getRootNode)
					.map(TreeEntry::getValue)
					.orElse(attribute);
			Assert.isTrue(
					ObjectUtil.isNull(resolvedAttribute)
							|| ClassUtil.isAssignable(attribute.getAttributeType(), resolvedAttribute.getAttributeType()),
					"return type of the root alias method [{}] is inconsistent with the original [{}]",
					resolvedAttribute.getClass(), attribute.getAttributeType()
			);
			if (attribute != resolvedAttribute) {
				attributeMap.put(attributeName, new ForceAliasedAnnotationAttribute(attribute, resolvedAttribute));
			}
		});
		synthesizedAnnotation.setAttributes(attributeMap);
	}

}
