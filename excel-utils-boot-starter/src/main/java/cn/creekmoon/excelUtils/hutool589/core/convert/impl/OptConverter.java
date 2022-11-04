package cn.creekmoon.excelUtils.hutool589.core.convert.impl;

import cn.creekmoon.excelUtils.hutool589.core.convert.AbstractConverter;
import cn.creekmoon.excelUtils.hutool589.core.lang.Opt;

/**
 * {@link Opt}对象转换器
 *
 * @author Looly
 * @since 5.7.16
 */
public class OptConverter extends AbstractConverter<Opt<?>> {
    private static final long serialVersionUID = 1L;

    @Override
    protected Opt<?> convertInternal(Object value) {
        return Opt.ofNullable(value);
    }

}
