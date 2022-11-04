package cn.creekmoon.excelUtils.hutool589.core.lang.generator;

import cn.creekmoon.excelUtils.hutool589.core.util.IdUtil;

/**
 * UUID生成器
 *
 * @author looly
 * @since 5.4.3
 */
public class UUIDGenerator implements Generator<String> {
	@Override
	public String next() {
		return IdUtil.fastUUID();
	}
}
