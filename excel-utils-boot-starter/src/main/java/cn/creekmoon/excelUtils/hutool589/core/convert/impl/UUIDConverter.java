package cn.creekmoon.excelUtils.hutool589.core.convert.impl;

import cn.creekmoon.excelUtils.hutool589.core.convert.AbstractConverter;

import java.util.UUID;

/**
 * UUID对象转换器转换器
 *
 * @author Looly
 * @since 4.0.10
 */
public class UUIDConverter extends AbstractConverter<UUID> {
	private static final long serialVersionUID = 1L;

	@Override
	protected UUID convertInternal(Object value) {
		return UUID.fromString(convertToStr(value));
	}

}
